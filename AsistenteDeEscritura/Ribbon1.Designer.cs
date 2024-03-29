﻿
namespace AsistenteDeEscritura
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Limpiar los recursos que se estén usando.
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
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.Rimas = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.Gerundios = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.Guiones = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.button3 = this.Factory.CreateRibbonButton();
            this.estadisticasButton = this.Factory.CreateRibbonToggleButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.Rimas);
            this.group1.Items.Add(this.button7);
            this.group1.Items.Add(this.separator3);
            this.group1.Items.Add(this.button4);
            this.group1.Items.Add(this.button5);
            this.group1.Items.Add(this.button6);
            this.group1.Items.Add(this.Gerundios);
            this.group1.Items.Add(this.separator2);
            this.group1.Items.Add(this.Guiones);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.button3);
            this.group1.Items.Add(this.estadisticasButton);
            this.group1.Label = "Asistente de Escritura";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.Image = global::AsistenteDeEscritura.Properties.Resources.writing2;
            this.button1.Label = "Repeticiones";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // Rimas
            // 
            this.Rimas.Image = global::AsistenteDeEscritura.Properties.Resources.poesia;
            this.Rimas.Label = "Rimas";
            this.Rimas.Name = "Rimas";
            this.Rimas.ShowImage = true;
            this.Rimas.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Rimas_Click);
            // 
            // button7
            // 
            this.button7.Image = global::AsistenteDeEscritura.Properties.Resources.megafono;
            this.button7.Label = "Cacofonía";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button7_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // button4
            // 
            this.button4.Image = global::AsistenteDeEscritura.Properties.Resources.marker;
            this.button4.Label = "Adv. mente";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Image = global::AsistenteDeEscritura.Properties.Resources.hablar;
            this.button5.Label = "Dicientes";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Image = global::AsistenteDeEscritura.Properties.Resources.paleta_de_pintura;
            this.button6.Label = "Adjetivos";
            this.button6.Name = "button6";
            this.button6.ShowImage = true;
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // Gerundios
            // 
            this.Gerundios.Image = global::AsistenteDeEscritura.Properties.Resources._263883;
            this.Gerundios.Label = "Gerundios";
            this.Gerundios.Name = "Gerundios";
            this.Gerundios.ShowImage = true;
            this.Gerundios.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Gerundios_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // Guiones
            // 
            this.Guiones.Image = global::AsistenteDeEscritura.Properties.Resources.conversation_icon;
            this.Guiones.Label = "Guiones";
            this.Guiones.Name = "Guiones";
            this.Guiones.ShowImage = true;
            this.Guiones.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Guiones_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = global::AsistenteDeEscritura.Properties.Resources.eraser;
            this.button3.Label = "Limpiar";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // estadisticasButton
            // 
            this.estadisticasButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.estadisticasButton.Image = global::AsistenteDeEscritura.Properties.Resources.lupa;
            this.estadisticasButton.Label = "Análisis";
            this.estadisticasButton.Name = "estadisticasButton";
            this.estadisticasButton.ShowImage = true;
            this.estadisticasButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.estadisticasButton_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Rimas;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Gerundios;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Guiones;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton estadisticasButton;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
