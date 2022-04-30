
namespace AsistenteDeEscritura
{
    partial class StatsPanel
    {
        /// <summary> 
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            System.Windows.Forms.ListViewGroup listViewGroup1 = new System.Windows.Forms.ListViewGroup("Raras", System.Windows.Forms.HorizontalAlignment.Left);
            System.Windows.Forms.ListViewGroup listViewGroup2 = new System.Windows.Forms.ListViewGroup("Comunes", System.Windows.Forms.HorizontalAlignment.Left);
            System.Windows.Forms.ListViewItem listViewItem1 = new System.Windows.Forms.ListViewItem(new string[] {
            "Bla",
            "Bla 1",
            "Bla 2",
            "Bla 3"}, -1);
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabUso = new System.Windows.Forms.TabPage();
            this.statsView = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.tabRitmo = new System.Windows.Forms.TabPage();
            this.ritmoView = new System.Windows.Forms.ListView();
            this.Frase = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Longitud = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Atomos = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.button1 = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabUso.SuspendLayout();
            this.tabRitmo.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabUso);
            this.tabControl1.Controls.Add(this.tabRitmo);
            this.tabControl1.Location = new System.Drawing.Point(6, 5);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(688, 392);
            this.tabControl1.TabIndex = 0;
            // 
            // tabUso
            // 
            this.tabUso.Controls.Add(this.statsView);
            this.tabUso.Location = new System.Drawing.Point(4, 29);
            this.tabUso.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabUso.Name = "tabUso";
            this.tabUso.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabUso.Size = new System.Drawing.Size(680, 359);
            this.tabUso.TabIndex = 0;
            this.tabUso.Text = "Repeticiones";
            this.tabUso.UseVisualStyleBackColor = true;
            // 
            // statsView
            // 
            this.statsView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.statsView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1});
            this.statsView.FullRowSelect = true;
            listViewGroup1.Header = "Raras";
            listViewGroup1.Name = "Raras";
            listViewGroup2.Header = "Comunes";
            listViewGroup2.Name = "Comunes";
            this.statsView.Groups.AddRange(new System.Windows.Forms.ListViewGroup[] {
            listViewGroup1,
            listViewGroup2});
            this.statsView.HideSelection = false;
            this.statsView.Location = new System.Drawing.Point(4, 5);
            this.statsView.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.statsView.MultiSelect = false;
            this.statsView.Name = "statsView";
            this.statsView.ShowItemToolTips = true;
            this.statsView.Size = new System.Drawing.Size(664, 345);
            this.statsView.TabIndex = 0;
            this.statsView.UseCompatibleStateImageBehavior = false;
            this.statsView.View = System.Windows.Forms.View.SmallIcon;
            this.statsView.SelectedIndexChanged += new System.EventHandler(this.listView1_SelectedIndexChanged);
            // 
            // tabRitmo
            // 
            this.tabRitmo.Controls.Add(this.ritmoView);
            this.tabRitmo.Location = new System.Drawing.Point(4, 29);
            this.tabRitmo.Name = "tabRitmo";
            this.tabRitmo.Size = new System.Drawing.Size(680, 359);
            this.tabRitmo.TabIndex = 1;
            this.tabRitmo.Text = "Frases";
            this.tabRitmo.UseVisualStyleBackColor = true;
            // 
            // ritmoView
            // 
            this.ritmoView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ritmoView.AutoArrange = false;
            this.ritmoView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.Frase,
            this.Longitud,
            this.Atomos});
            this.ritmoView.FullRowSelect = true;
            this.ritmoView.GridLines = true;
            this.ritmoView.HideSelection = false;
            this.ritmoView.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
            listViewItem1});
            this.ritmoView.Location = new System.Drawing.Point(3, 3);
            this.ritmoView.MultiSelect = false;
            this.ritmoView.Name = "ritmoView";
            this.ritmoView.ShowGroups = false;
            this.ritmoView.ShowItemToolTips = true;
            this.ritmoView.Size = new System.Drawing.Size(674, 353);
            this.ritmoView.TabIndex = 0;
            this.ritmoView.UseCompatibleStateImageBehavior = false;
            this.ritmoView.View = System.Windows.Forms.View.Details;
            this.ritmoView.DoubleClick += new System.EventHandler(this.ritmoView_DoubleClick);
            // 
            // Frase
            // 
            this.Frase.Text = "Frase";
            this.Frase.Width = 150;
            // 
            // Longitud
            // 
            this.Longitud.Text = "Longitud";
            this.Longitud.Width = 152;
            // 
            // Atomos
            // 
            this.Atomos.Text = "Ritmos";
            this.Atomos.Width = 198;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button1.BackgroundImage = global::AsistenteDeEscritura.Properties.Resources.refresh_;
            this.button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.Location = new System.Drawing.Point(6, 400);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button1.Name = "button1";
            this.button1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.button1.Size = new System.Drawing.Size(82, 68);
            this.button1.TabIndex = 1;
            this.button1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // StatsPanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.tabControl1);
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "StatsPanel";
            this.Size = new System.Drawing.Size(698, 478);
            this.Load += new System.EventHandler(this.StatsPanel_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabUso.ResumeLayout(false);
            this.tabRitmo.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabUso;
        private System.Windows.Forms.ListView statsView;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.TabPage tabRitmo;
        private System.Windows.Forms.ListView ritmoView;
        private System.Windows.Forms.ColumnHeader Frase;
        private System.Windows.Forms.ColumnHeader Longitud;
        private System.Windows.Forms.ColumnHeader Atomos;
    }
}
