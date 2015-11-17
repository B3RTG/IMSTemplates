namespace IMSRerports_ImportService
{
    partial class ImportService
    {
        /// <summary> 
        /// Variable del diseñador requerida.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        //private System.Diagnostics.EventLog ImportService_EventLog;
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
            this.ImportServices_EventLog = new System.Diagnostics.EventLog();
            this.AutoLog = false;
            ((System.ComponentModel.ISupportInitialize)(this.ImportServices_EventLog)).BeginInit();
            // 
            // ImportService
            // 
            this.ServiceName = "ImportServices";
            ((System.ComponentModel.ISupportInitialize)(this.ImportServices_EventLog)).EndInit();

        }

        #endregion

        private System.Diagnostics.EventLog ImportServices_EventLog;
    }
}
