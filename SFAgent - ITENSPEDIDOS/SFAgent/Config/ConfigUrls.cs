namespace SFAgent.Config
{
    public static class ConfigUrls
    {
        public static string AuthUrl =
            "https://acos-continente--homolog.sandbox.my.salesforce.com/services/oauth2/token";

        // Base do sObject usado no upsert (mantém como está)
        public static string ApiCondicaoBase =
            "https://acos-continente--homolog.sandbox.my.salesforce.com/services/data/v60.0/sobjects/OrderItem";

        // Campo External Id usado no upsert
        public static string ApiCondicaoExternalField = "CA_IdExterno__c";

        // (Opcional) REST base explícita; se preencher, a API usa isso para /query e /sobjects
        // Ex.: "https://acos-continente--homolog.sandbox.my.salesforce.com/services/data/v60.0"
        public static string ApiRestBase = null;
    }
}
