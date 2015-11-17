namespace IMSRerports_ImportService
{
    partial class ProjectInstaller
    {
        /// <summary>
        /// Variable del diseñador requerida.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Limpiar los recursos que se estén utilizando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de componentes

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.IMSImportService_InstalerProc = new System.ServiceProcess.ServiceProcessInstaller();
            this.IMSImportService_Installer = new System.ServiceProcess.ServiceInstaller();
            // 
            // IMSImportService_InstalerProc
            // 
            this.IMSImportService_InstalerProc.Account = System.ServiceProcess.ServiceAccount.LocalService;
            this.IMSImportService_InstalerProc.Password = null;
            this.IMSImportService_InstalerProc.Username = null;
            // 
            // IMSImportService_Installer
            // 
            this.IMSImportService_Installer.Description = "Service for import & format data of reports for IMS";
            this.IMSImportService_Installer.DisplayName = "IMSImportService";
            this.IMSImportService_Installer.ServiceName = "ImportServices";
            this.IMSImportService_Installer.StartType = System.ServiceProcess.ServiceStartMode.Automatic;
            this.IMSImportService_Installer.AfterInstall += new System.Configuration.Install.InstallEventHandler(this.IMSImportService_Installer_AfterInstall);
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.IMSImportService_InstalerProc,
            this.IMSImportService_Installer});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller IMSImportService_InstalerProc;
        private System.ServiceProcess.ServiceInstaller IMSImportService_Installer;
    }
}