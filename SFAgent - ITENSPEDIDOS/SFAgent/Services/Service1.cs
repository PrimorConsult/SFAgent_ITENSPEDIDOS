using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SFAgent.Salesforce;
using SFAgent.Sap;
using SFAgent.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.ServiceProcess;
using System.Text; // <- para normalização sem acento
using System.Threading;
using System.Threading.Tasks;

namespace SFAgent.Services
{
    public partial class Service1 : ServiceBase
    {
        private SalesforceAuth _auth;
        private SalesforceApi _api;
        private Timer _timer;

        // Se tiver uma Pricebook2 com ExternalId, coloque aqui (senão deixe null)
        private const string DEFAULT_PRICEBOOK_EXT = null; // ex.: "PB-ACOS-DEFAULT"

        // LIGADO: define automaticamente a Pricebook no Pedido (se regra de validação permitir)
        private const bool ENABLE_AUTO_SET_PRICEBOOK = true;

        public Service1()
        {
            InitializeComponent();
            _auth = new SalesforceAuth();
            _api = new SalesforceApi();
        }

        protected override void OnStart(string[] args)
        {
            if (!System.Diagnostics.EventLog.SourceExists("SFAgent"))
                System.Diagnostics.EventLog.CreateEventSource("SFAgent", "Application");

            Task.Run(async () =>
            {
                try
                {
                    await ProcessarItens();
                }
                catch (Exception ex)
                {
                    Logger.Log($"Erro inicial no OnStart (ItensPedidos): {ex.Message}");
                }
            });

            _timer = new Timer(async _ =>
            {
                try
                {
                    await ProcessarItens();
                }
                catch (Exception ex)
                {
                    Logger.Log($"Erro no Timer (ItensPedidos): {ex.Message}");
                }
            }, null, TimeSpan.FromMinutes(5), TimeSpan.FromMinutes(5));
        }

        protected override void OnStop()
        {
            _timer?.Dispose();
            Logger.Log("Serviço parado (ItensPedidos).");
        }

        // ---------- Helpers ----------
        private static bool IsDbNull(object v) => v == null || v == DBNull.Value;
        private static string S(object v) => IsDbNull(v) ? null : v.ToString();

        private static string AsDate(object dt)
        {
            if (IsDbNull(dt)) return null;
            var d = Convert.ToDateTime(dt, CultureInfo.InvariantCulture);
            return d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        private static int? I(object v)
        {
            if (IsDbNull(v)) return null;
            try { return Convert.ToInt32(v, CultureInfo.InvariantCulture); } catch { return null; }
        }

        private static decimal? M2(object v)
        {
            if (IsDbNull(v)) return null;
            try
            {
                var d = Convert.ToDecimal(v, CultureInfo.InvariantCulture);
                return Math.Round(d, 2, MidpointRounding.AwayFromZero);
            }
            catch { return null; }
        }

        private static decimal? M3(object v)
        {
            if (IsDbNull(v)) return null;
            try
            {
                var d = Convert.ToDecimal(v, CultureInfo.InvariantCulture);
                return Math.Round(d, 3, MidpointRounding.AwayFromZero);
            }
            catch { return null; }
        }

        private static string Trunc(string s, int max)
            => string.IsNullOrEmpty(s) ? s : (s.Length <= max ? s : s.Substring(0, max));

        private static bool Yes(object v)
        {
            var s = S(v);
            return !string.IsNullOrWhiteSpace(s) &&
                   (s.Equals("Y", StringComparison.OrdinalIgnoreCase) ||
                    s == "1" ||
                    s.Equals("T", StringComparison.OrdinalIgnoreCase));
        }

        private static object Lookup(string ext) => string.IsNullOrWhiteSpace(ext) ? null : new { CA_IdExterno__c = ext };
        private static object Lookup(int? ext) => !ext.HasValue ? null : new { CA_IdExterno__c = ext.Value.ToString(CultureInfo.InvariantCulture) };

        private static string MapMotivoAlteracao(string motivoRaw)
        {
            if (string.IsNullOrWhiteSpace(motivoRaw)) return null;
            var m = motivoRaw.Trim().ToLowerInvariant();
            if (m.Contains("condi") && m.Contains("pag")) return "Condição de Pagamento";
            if (m.Contains("crédit") || m.Contains("credito")) return "Crédito";
            if (m.Contains("estoq")) return "Estoque";
            if (m.Contains("prazo")) return "Prazo de Entrega";
            if (m.Contains("preç") || m.Contains("preco") || m.Contains("preço")) return "Preço";
            if (m.Contains("residual")) return "Quantidade Residual";
            if (m.Contains("tomada")) return "Tomada de Preço";
            if (m.Contains("troca")) return "Troca de Pedido";
            return null;
        }

        private static string SoqlEsc(string v)
        {
            if (string.IsNullOrEmpty(v)) return v;
            return v.Replace("\\", "\\\\").Replace("'", "\\'");
        }

        // normalização p/ comparar strings sem acento e case-insensitive
        private static string NormalizeKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            var formD = s.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(formD.Length);
            foreach (var ch in formD)
            {
                var uc = CharUnicodeInfo.GetUnicodeCategory(ch);
                if (uc != UnicodeCategory.NonSpacingMark) sb.Append(ch);
            }
            return sb.ToString().ToLowerInvariant().Trim();
        }

        private static bool EqualsAnyStatus(string status, params string[] options)
        {
            var ns = NormalizeKey(status);
            for (int i = 0; i < options.Length; i++)
            {
                if (ns == NormalizeKey(options[i])) return true;
            }
            return false;
        }
        // ---------- /Helpers ----------

        private async Task ProcessarItens()
        {
            try
            {
                var token = await _auth.GetValidToken(); // Bearer token (string)

                // Ajuste conforme ambiente
                var sap = new SapConnector("HANADB:30015", "SBO_ACOS_TESTE", "B1ADMIN", "S4P@2Q60_tm2");

                var sql = @"CALL SP_ITENSPEDIDO_SF();";
                var rows = sap.ExecuteQuery(sql);

                int insertCount = 0;
                int updateCount = 0;
                int errorCount = 0;

                foreach (var r in rows)
                {
                    SalesforceApi.UpsertResult result = null;
                    var docNum = S(r["DocNum"]);
                    var lineNum = I(r["LineNum"]) ?? 0;
                    var itemExternalId = S(r["IdExternoItem"]);

                    if (string.IsNullOrWhiteSpace(itemExternalId))
                    {
                        Logger.Log($"Item ignorado: IdExternoItem vazio (DocNum={docNum}, LineNum={lineNum}).");
                        continue;
                    }

                    try
                    {
                        // --- Relacionamentos / chaves ---
                        var itemCode = Trunc(S(r["ItemCode"]), 50);
                        var whsCode = Trunc(S(r["WhsCode"]), 8);
                        var taxCode = Trunc(S(r["TaxCode"]), 8);

                        // --- Numéricos / texto ---
                        var quantity = M3(r["Quantity"]);
                        var unitPrice = M2(r["UnitPrice"]);
                        var lineTotal = M2(r["LineTotal"]);
                        var markup = M2(r["Markup"]);
                        var qtdAuthPend = Trunc(S(r["QtdAuthPend"]), 255);
                        var qtdUnFat = M3(r["QtdUnFat"]);
                        var unFat = Trunc(S(r["UnFat"]), 10);
                        var comissao = M2(r["Comissao"]);
                        var freeCharge = Yes(r["GratisYN"]);

                        // --- Tributos / componentes ---
                        var vICMS = M2(r["VlrICMS"]);
                        var vPIS = M2(r["VlrPIS"]);
                        var vCOFINS = M2(r["VlrCOFINS"]);
                        var vTxFin = M2(r["VlrTaxaFin"]);

                        // --- Datas e outros ---
                        var dtEntrega = AsDate(r["DataEntrega"]);
                        var motivoAlt = MapMotivoAlteracao(S(r["MotivoAlteracao"]));
                        var observacao = Trunc(S(r["Observacao"]), 255);
                        var pesoLiq = M3(r["PesoLiquido"]);

                        // --- Fiscais ---
                        var cfop = Trunc(S(r["CFOP"]), 6);
                        var cst = Trunc(S(r["CST"]), 6);

                        // --- Relacionamentos custom opcionais ---
                        var usoPrincipalLookup = Lookup(I(r["Usage"]));
                        var depositoLookup = Lookup(whsCode);

                        if (string.IsNullOrWhiteSpace(docNum))
                        {
                            Logger.Log($"Item ignorado: DocNum vazio (LineNum={lineNum}).");
                            continue;
                        }
                        if (string.IsNullOrWhiteSpace(itemCode))
                        {
                            Logger.Log($"Item ignorado: ItemCode vazio (DocNum={docNum}, LineNum={lineNum}).");
                            continue;
                        }

                        // ---------- (1) Buscar o Pedido ----------
                        var docNumEsc = SoqlEsc(docNum);
                        var order = await _api.QuerySingleAsync(token,
                            $"SELECT Id, Pricebook2Id, Status, ActivatedDate " +
                            $"FROM Order WHERE CA_IdExterno__c = '{docNumEsc}' LIMIT 1");

                        if (order == null)
                        {
                            Logger.Log($"Item ignorado: Pedido não encontrado no SF (DocNum={docNum}).");
                            continue;
                        }

                        var status = order.Value<string>("Status");
                        var activatedDt = order.Value<DateTime?>("ActivatedDate");

                        // Multi-idioma: considerar rótulos/valores em PT/EN
                        bool statusIsDraft = EqualsAnyStatus(status, "Draft", "Rascunho");
                        bool statusIsActivated = EqualsAnyStatus(status, "Activated", "Ativado", "Activado");

                        // Pedido é editável apenas se for "Draft/Rascunho" e NÃO tiver ActivatedDate
                        bool isEditable = statusIsDraft && !activatedDt.HasValue;

                        if (!isEditable)
                        {
                            var st = string.IsNullOrWhiteSpace(status) ? "desconhecido" : status;
                            var reason = (statusIsActivated || activatedDt.HasValue) ? "ativado" : "status != Draft";
                            Logger.Log($"Item ignorado: Pedido não editável ({reason}, Status={st}) (DocNum={docNum}).");
                            continue;
                        }

                        var pricebook2Id = order.Value<string>("Pricebook2Id");
                        var orderId = order.Value<string>("Id");

                        // ---------- (1.1) AUTO-DEFINIÇÃO DE PRICEBOOK (com múltiplos fallbacks) ----------
                        if (string.IsNullOrEmpty(pricebook2Id))
                        {
                            if (!ENABLE_AUTO_SET_PRICEBOOK)
                            {
                                Logger.Log($"Item ignorado: Pedido sem Pricebook e auto-definição desativada por configuração (DocNum={docNum}).");
                                continue;
                            }

                            string chosenPbId = null;

                            // 1) Preferir uma Pricebook configurada por ExternalId (se existir em Pricebook2)
                            if (!string.IsNullOrEmpty(DEFAULT_PRICEBOOK_EXT))
                            {
                                var pbCfg = await _api.QuerySingleAsync(token,
                                    $"SELECT Id, IsActive FROM Pricebook2 WHERE CA_IdExterno__c = '{SoqlEsc(DEFAULT_PRICEBOOK_EXT)}' LIMIT 1");

                                if (pbCfg != null)
                                {
                                    chosenPbId = pbCfg.Value<string>("Id");
                                    var cfgActive = pbCfg.Value<bool?>("IsActive") == true;
                                    if (!cfgActive)
                                    {
                                        try
                                        {
                                            await _api.PatchSObject(token, "Pricebook2", chosenPbId, new { IsActive = true });
                                        }
                                        catch (Exception exActCfg)
                                        {
                                            Logger.Log($"Aviso: Falha ao ativar Pricebook '{DEFAULT_PRICEBOOK_EXT}': {exActCfg.Message}");
                                        }
                                    }
                                }
                                else
                                {
                                    Logger.Log($"Aviso: Pricebook por ExternalId '{DEFAULT_PRICEBOOK_EXT}' não encontrada.");
                                }
                            }

                            // 2) Tentar a Standard (ativando se necessário)
                            if (chosenPbId == null)
                            {
                                var stdPb = await _api.QuerySingleAsync(token,
                                    "SELECT Id, IsActive FROM Pricebook2 WHERE IsStandard = true LIMIT 1");

                                if (stdPb != null)
                                {
                                    chosenPbId = stdPb.Value<string>("Id");
                                    var stdActive = stdPb.Value<bool?>("IsActive") == true;

                                    if (!stdActive)
                                    {
                                        try
                                        {
                                            await _api.PatchSObject(token, "Pricebook2", chosenPbId, new { IsActive = true });
                                            Logger.Log($"Standard Price Book ativada automaticamente (Id={chosenPbId}).");
                                        }
                                        catch (Exception exStd)
                                        {
                                            Logger.Log($"Aviso: Não foi possível ativar a Standard Price Book: {exStd.Message}");
                                        }
                                    }
                                }
                                else
                                {
                                    Logger.Log("Aviso: Nenhuma Standard Price Book encontrada no org.");
                                }
                            }

                            // 3) Qualquer Pricebook ativa
                            if (chosenPbId == null)
                            {
                                var anyActivePb = await _api.QuerySingleAsync(token,
                                    "SELECT Id FROM Pricebook2 WHERE IsActive = true ORDER BY IsStandard DESC, CreatedDate ASC LIMIT 1");

                                if (anyActivePb != null)
                                {
                                    chosenPbId = anyActivePb.Value<string>("Id");
                                    Logger.Log($"Usando Pricebook ativa encontrada (Id={chosenPbId}).");
                                }
                            }

                            if (chosenPbId == null)
                            {
                                Logger.Log($"Item ignorado: Pedido sem Pricebook e nenhuma Pricebook ativa configurada/encontrada (DocNum={docNum}).");
                                continue;
                            }

                            try
                            {
                                await _api.PatchSObject(token, "Order", orderId, new { Pricebook2Id = chosenPbId });
                                pricebook2Id = chosenPbId;
                                Logger.Log($"Pedido {docNum}: Pricebook definida automaticamente ({chosenPbId}).");
                            }
                            catch (Exception exSetPb)
                            {
                                var msg = exSetPb.Message ?? string.Empty;
                                if (msg.Contains("FIELD_CUSTOM_VALIDATION_EXCEPTION") ||
                                    msg.Contains("Você só pode editar"))
                                {
                                    Logger.Log($"Item ignorado: Regra de validação bloqueou a edição da Pricebook no pedido {docNum}. " +
                                               $"Peça ao admin para liberar a integração (custom permission) ou permitir setar Pricebook2Id quando vazio. " +
                                               $"Detalhe: {exSetPb.Message}");
                                }
                                else
                                {
                                    Logger.Log($"Item ignorado: Falha ao definir Pricebook no pedido {docNum}: {exSetPb.Message}");
                                }
                                continue;
                            }
                        }

                        // ---------- (2) Buscar o PricebookEntry do produto nessa Pricebook ----------
                        var itemCodeEsc = SoqlEsc(itemCode);
                        var pbe = await _api.QuerySingleAsync(token,
                            "SELECT Id, UnitPrice, IsActive " +
                            "FROM PricebookEntry " +
                            $"WHERE Pricebook2Id = '{pricebook2Id}' " +
                            $"AND Product2.CA_IdExterno__c = '{itemCodeEsc}' " +
                            "AND IsActive = true " +
                            "LIMIT 1");

                        if (pbe == null)
                        {
                            Logger.Log($"Item ignorado: Produto {itemCode} não está na Pricebook do pedido (DocNum={docNum}).");
                            continue;
                        }

                        var pricebookEntryId = pbe.Value<string>("Id");

                        // ---------- (3) Ajuste UnitPrice (se necessário) ----------
                        if (!unitPrice.HasValue && lineTotal.HasValue && quantity.HasValue && quantity.Value != 0)
                            unitPrice = Math.Round(lineTotal.Value / quantity.Value, 2, MidpointRounding.AwayFromZero);

                        // ---------- (4) Body do OrderItem (sem TotalPrice, sem Product2) ----------
                        var body = new
                        {
                            // Upsert via External Id NA URL (não mandar CA_IdExterno__c do próprio OrderItem no body)
                            OrderId = orderId,
                            PricebookEntryId = pricebookEntryId,

                            Quantity = quantity,
                            UnitPrice = unitPrice,

                            // Seus campos custom
                            CA_CodImposto__c = taxCode,
                            CA_QtdAuthPend__c = qtdAuthPend,
                            CA_UnFat__c = unFat,

                            CA_Markup__c = markup,
                            CA_QtdUnFat__c = qtdUnFat,
                            CA_Comissao__c = comissao,
                            CA_Gratuito__c = freeCharge,

                            CA_ValorICMS__c = vICMS,
                            CA_ValorPIS__c = vPIS,
                            CA_ValorCOFINS__c = vCOFINS,
                            CA_TaxaFinanceira__c = vTxFin,
                            CA_PesoLiquido__c = pesoLiq,

                            CA_DataEntrega__c = dtEntrega,
                            CA_MotivoAlteracao__c = motivoAlt,
                            CA_Observacao__c = observacao,

                            CA_CFOP__c = cfop,
                            CA_CST__c = cst,

                            // integração
                            CA_StatusIntegracao__c = "Integrado",
                            CA_RetornoIntegracao__c = "Integrado",
                            CA_AtualizacaoERP__c = DateTime.UtcNow.ToString("s") + "Z",

                            // Relacionamentos custom opcionais
                            CA_Deposito__r = depositoLookup,
                            CA_UsoPrincipal__r = usoPrincipalLookup
                        };

                        // ---------- (5) Upsert via URL (ExternalId na URL; NUNCA no body) ----------

                        if (await _api.UpsertItemPedido(token, itemExternalId, body) != null && !string.IsNullOrEmpty((await _api.UpsertItemPedido(token, itemExternalId, body)).Outcome))
                        {
                            var up = (await _api.UpsertItemPedido(token, itemExternalId, body)).Outcome.ToUpperInvariant();
                            if (up == "POST" || up == "INSERT") insertCount++;
                            else if (up == "PATCH" || up == "UPDATE") updateCount++;
                        }

                        Logger.Log($"SF Itens {(await _api.UpsertItemPedido(token, itemExternalId, body))?.Outcome} | METHOD={(await _api.UpsertItemPedido(token, itemExternalId, body))?.Method} | ExternalId={itemExternalId} | Status={(await _api.UpsertItemPedido(token, itemExternalId, body))?.StatusCode}");
                    }
                    catch (Exception exItem)
                    {
                        errorCount++;
                        var rowJson = JsonConvert.SerializeObject(r);
                        Logger.Log(
                            $"ERRO Itens | ExternalId={itemExternalId} | Msg={exItem.Message} | Row={rowJson}",
                            asError: true
                        );
                    }
                }

                Logger.Log($"Sync Itens finalizado. | Inseridos={insertCount} | Atualizados={updateCount} | Erros={errorCount} | Total Processados={rows.Count()}.");
            }
            catch (Exception ex)
            {
                Logger.Log($"Erro ao processar itens: {ex.Message}");
            }
        }
    }
}
