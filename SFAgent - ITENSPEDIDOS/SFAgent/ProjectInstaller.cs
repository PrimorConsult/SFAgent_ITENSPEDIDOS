using System.ComponentModel;
using System.ServiceProcess;

namespace SFAgent
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        private ServiceProcessInstaller _processInstaller;
        private ServiceInstaller _serviceInstaller;

        public ProjectInstaller()
        {
            _processInstaller = new ServiceProcessInstaller();
            _serviceInstaller = new ServiceInstaller();

            _processInstaller.Account = ServiceAccount.LocalSystem;

            _serviceInstaller.ServiceName = "SFAgent - ItensPedidos";
            _serviceInstaller.DisplayName = "SFAgent - ItensPedidos";
            _serviceInstaller.StartType = ServiceStartMode.Automatic;
            _serviceInstaller.Description = "Serviço que integra Itens do Pedidos do SAP HANA (RDR1) com Salesforce";

            Installers.Add(_processInstaller);
            Installers.Add(_serviceInstaller);
        }
    }
}