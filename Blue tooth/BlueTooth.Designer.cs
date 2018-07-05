namespace Blue_tooth
{
    partial class BlueTooth
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BlueTooth));
            this.label3 = new System.Windows.Forms.Label();
            this.Att = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.GPIB = new System.Windows.Forms.TextBox();
            this.Receiver = new System.Windows.Forms.TabPage();
            this.BERItem = new System.Windows.Forms.GroupBox();
            this.RCVMaximumInputLevel = new System.Windows.Forms.CheckBox();
            this.EDRsensitivity = new System.Windows.Forms.CheckBox();
            this.RCVsensitivityMeasurements = new System.Windows.Forms.CheckBox();
            this.EDRmaximunInputLevel = new System.Windows.Forms.CheckBox();
            this.EDRBERfloorPerformance = new System.Windows.Forms.CheckBox();
            this.splitter1 = new System.Windows.Forms.Splitter();
            this.Transmitter = new System.Windows.Forms.TabPage();
            this.LE = new System.Windows.Forms.CheckBox();
            this.EDR = new System.Windows.Forms.CheckBox();
            this.BR = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.MaxRelativeTransmit = new System.Windows.Forms.CheckBox();
            this.InbandSpuriousEmissions = new System.Windows.Forms.CheckBox();
            this.DifferentialPhaseEncoding = new System.Windows.Forms.CheckBox();
            this.EDRCarrierFrequency = new System.Windows.Forms.CheckBox();
            this.MinRelativeTransmit = new System.Windows.Forms.CheckBox();
            this.EDRTransmitter = new System.Windows.Forms.GroupBox();
            this.label12 = new System.Windows.Forms.Label();
            this.T3DH5 = new System.Windows.Forms.CheckBox();
            this.T2DH1 = new System.Windows.Forms.CheckBox();
            this.T3DH3 = new System.Windows.Forms.CheckBox();
            this.T2DH3 = new System.Windows.Forms.CheckBox();
            this.T3DH1 = new System.Windows.Forms.CheckBox();
            this.T2DH5 = new System.Windows.Forms.CheckBox();
            this.TransmitterTest = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.T1DH3 = new System.Windows.Forms.CheckBox();
            this.T1DH5 = new System.Windows.Forms.CheckBox();
            this.T1DH1 = new System.Windows.Forms.CheckBox();
            this.TransmitterSetting = new System.Windows.Forms.GroupBox();
            this.TRMcarrierFrequency = new System.Windows.Forms.CheckBox();
            this.TRMfrequenyRange = new System.Windows.Forms.CheckBox();
            this.TRM20dBBandwidth = new System.Windows.Forms.CheckBox();
            this.TRMinitialCarrier = new System.Windows.Forms.CheckBox();
            this.TRMpowerControl = new System.Windows.Forms.CheckBox();
            this.TRMadjacentChannel = new System.Windows.Forms.CheckBox();
            this.TRMoutputPower = new System.Windows.Forms.CheckBox();
            this.TRMpowerDensity = new System.Windows.Forms.CheckBox();
            this.TRMmodulationCharacteristics = new System.Windows.Forms.CheckBox();
            this.Sett = new System.Windows.Forms.TabControl();
            this.MaxTransmintPower = new System.Windows.Forms.CheckBox();
            this.MinTransmitPower = new System.Windows.Forms.CheckBox();
            this.StartTest = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.CMWinfo = new System.Windows.Forms.Label();
            this.Port = new System.Windows.Forms.ComboBox();
            this.NunCh = new System.Windows.Forms.TrackBar();
            this.Channellist = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.EventLog = new System.Windows.Forms.TextBox();
            this.ResultLog = new System.Windows.Forms.TextBox();
            this.Receiver.SuspendLayout();
            this.BERItem.SuspendLayout();
            this.Transmitter.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.EDRTransmitter.SuspendLayout();
            this.TransmitterTest.SuspendLayout();
            this.TransmitterSetting.SuspendLayout();
            this.Sett.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.NunCh)).BeginInit();
            this.SuspendLayout();
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.DarkOrange;
            this.label3.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(395, 74);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(50, 25);
            this.label3.TabIndex = 9;
            this.label3.Text = "(dB)";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Att
            // 
            this.Att.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Att.Location = new System.Drawing.Point(289, 74);
            this.Att.Name = "Att";
            this.Att.Size = new System.Drawing.Size(100, 25);
            this.Att.TabIndex = 8;
            this.Att.Text = "2.5";
            // 
            // label2
            // 
            this.label2.AccessibleRole = System.Windows.Forms.AccessibleRole.TitleBar;
            this.label2.BackColor = System.Drawing.Color.DarkOrange;
            this.label2.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(211, 74);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(50, 25);
            this.label2.TabIndex = 7;
            this.label2.Text = "ATT";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.DarkOrange;
            this.label1.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label1.Location = new System.Drawing.Point(211, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(50, 25);
            this.label1.TabIndex = 6;
            this.label1.Text = "GPIB";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // GPIB
            // 
            this.GPIB.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GPIB.Location = new System.Drawing.Point(289, 28);
            this.GPIB.Name = "GPIB";
            this.GPIB.Size = new System.Drawing.Size(100, 25);
            this.GPIB.TabIndex = 5;
            this.GPIB.Text = "20";
            // 
            // Receiver
            // 
            this.Receiver.BackColor = System.Drawing.Color.DarkGray;
            this.Receiver.Controls.Add(this.BERItem);
            this.Receiver.Controls.Add(this.splitter1);
            this.Receiver.Location = new System.Drawing.Point(4, 27);
            this.Receiver.Name = "Receiver";
            this.Receiver.Padding = new System.Windows.Forms.Padding(3);
            this.Receiver.Size = new System.Drawing.Size(839, 585);
            this.Receiver.TabIndex = 3;
            this.Receiver.Text = "Receiver";
            // 
            // BERItem
            // 
            this.BERItem.Controls.Add(this.RCVMaximumInputLevel);
            this.BERItem.Controls.Add(this.EDRsensitivity);
            this.BERItem.Controls.Add(this.RCVsensitivityMeasurements);
            this.BERItem.Controls.Add(this.EDRmaximunInputLevel);
            this.BERItem.Controls.Add(this.EDRBERfloorPerformance);
            this.BERItem.Location = new System.Drawing.Point(46, 40);
            this.BERItem.Name = "BERItem";
            this.BERItem.Size = new System.Drawing.Size(478, 280);
            this.BERItem.TabIndex = 8;
            this.BERItem.TabStop = false;
            this.BERItem.Text = "TestItem";
            // 
            // RCVMaximumInputLevel
            // 
            this.RCVMaximumInputLevel.AutoSize = true;
            this.RCVMaximumInputLevel.Checked = true;
            this.RCVMaximumInputLevel.CheckState = System.Windows.Forms.CheckState.Checked;
            this.RCVMaximumInputLevel.Location = new System.Drawing.Point(30, 100);
            this.RCVMaximumInputLevel.Name = "RCVMaximumInputLevel";
            this.RCVMaximumInputLevel.Size = new System.Drawing.Size(262, 22);
            this.RCVMaximumInputLevel.TabIndex = 10;
            this.RCVMaximumInputLevel.Text = "RCV/CA/06/C-MaximumInputLevel";
            this.RCVMaximumInputLevel.UseVisualStyleBackColor = true;
            // 
            // EDRsensitivity
            // 
            this.EDRsensitivity.AutoSize = true;
            this.EDRsensitivity.Checked = true;
            this.EDRsensitivity.CheckState = System.Windows.Forms.CheckState.Checked;
            this.EDRsensitivity.Location = new System.Drawing.Point(30, 144);
            this.EDRsensitivity.Name = "EDRsensitivity";
            this.EDRsensitivity.Size = new System.Drawing.Size(246, 22);
            this.EDRsensitivity.TabIndex = 5;
            this.EDRsensitivity.Text = "RCV/CA/07/C-EDR Sensitivity";
            this.EDRsensitivity.UseVisualStyleBackColor = true;
            // 
            // RCVsensitivityMeasurements
            // 
            this.RCVsensitivityMeasurements.AutoSize = true;
            this.RCVsensitivityMeasurements.Checked = true;
            this.RCVsensitivityMeasurements.CheckState = System.Windows.Forms.CheckState.Checked;
            this.RCVsensitivityMeasurements.Location = new System.Drawing.Point(30, 60);
            this.RCVsensitivityMeasurements.Name = "RCVsensitivityMeasurements";
            this.RCVsensitivityMeasurements.Size = new System.Drawing.Size(342, 22);
            this.RCVsensitivityMeasurements.TabIndex = 1;
            this.RCVsensitivityMeasurements.Text = "RCV/CA/01-02/C-Sensitivity measurements";
            this.RCVsensitivityMeasurements.UseVisualStyleBackColor = true;
            // 
            // EDRmaximunInputLevel
            // 
            this.EDRmaximunInputLevel.AutoSize = true;
            this.EDRmaximunInputLevel.Checked = true;
            this.EDRmaximunInputLevel.CheckState = System.Windows.Forms.CheckState.Checked;
            this.EDRmaximunInputLevel.Location = new System.Drawing.Point(30, 227);
            this.EDRmaximunInputLevel.Name = "EDRmaximunInputLevel";
            this.EDRmaximunInputLevel.Size = new System.Drawing.Size(310, 22);
            this.EDRmaximunInputLevel.TabIndex = 9;
            this.EDRmaximunInputLevel.Text = "RCV/CA/10/C-EDR Maximun Input Level";
            this.EDRmaximunInputLevel.UseVisualStyleBackColor = true;
            // 
            // EDRBERfloorPerformance
            // 
            this.EDRBERfloorPerformance.AutoSize = true;
            this.EDRBERfloorPerformance.Checked = true;
            this.EDRBERfloorPerformance.CheckState = System.Windows.Forms.CheckState.Checked;
            this.EDRBERfloorPerformance.Location = new System.Drawing.Point(30, 188);
            this.EDRBERfloorPerformance.Name = "EDRBERfloorPerformance";
            this.EDRBERfloorPerformance.Size = new System.Drawing.Size(326, 22);
            this.EDRBERfloorPerformance.TabIndex = 7;
            this.EDRBERfloorPerformance.Text = "RCV/CA/08/C-EDR BER Floor Performance";
            this.EDRBERfloorPerformance.UseVisualStyleBackColor = true;
            // 
            // splitter1
            // 
            this.splitter1.Location = new System.Drawing.Point(3, 3);
            this.splitter1.Name = "splitter1";
            this.splitter1.Size = new System.Drawing.Size(3, 579);
            this.splitter1.TabIndex = 3;
            this.splitter1.TabStop = false;
            // 
            // Transmitter
            // 
            this.Transmitter.BackColor = System.Drawing.Color.DarkGray;
            this.Transmitter.Controls.Add(this.LE);
            this.Transmitter.Controls.Add(this.EDR);
            this.Transmitter.Controls.Add(this.BR);
            this.Transmitter.Controls.Add(this.groupBox3);
            this.Transmitter.Controls.Add(this.EDRTransmitter);
            this.Transmitter.Controls.Add(this.TransmitterTest);
            this.Transmitter.Controls.Add(this.TransmitterSetting);
            this.Transmitter.Location = new System.Drawing.Point(4, 27);
            this.Transmitter.Name = "Transmitter";
            this.Transmitter.Padding = new System.Windows.Forms.Padding(3);
            this.Transmitter.Size = new System.Drawing.Size(839, 585);
            this.Transmitter.TabIndex = 1;
            this.Transmitter.Text = "Transmitter";
            // 
            // LE
            // 
            this.LE.AutoSize = true;
            this.LE.BackColor = System.Drawing.Color.WhiteSmoke;
            this.LE.Location = new System.Drawing.Point(307, 50);
            this.LE.Name = "LE";
            this.LE.Size = new System.Drawing.Size(46, 22);
            this.LE.TabIndex = 25;
            this.LE.Text = "LE";
            this.LE.UseVisualStyleBackColor = false;
            // 
            // EDR
            // 
            this.EDR.AutoSize = true;
            this.EDR.BackColor = System.Drawing.Color.WhiteSmoke;
            this.EDR.Checked = true;
            this.EDR.CheckState = System.Windows.Forms.CheckState.Checked;
            this.EDR.Location = new System.Drawing.Point(164, 50);
            this.EDR.Name = "EDR";
            this.EDR.Size = new System.Drawing.Size(54, 22);
            this.EDR.TabIndex = 24;
            this.EDR.Text = "EDR";
            this.EDR.UseVisualStyleBackColor = false;
            // 
            // BR
            // 
            this.BR.AutoSize = true;
            this.BR.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.BR.Checked = true;
            this.BR.CheckState = System.Windows.Forms.CheckState.Checked;
            this.BR.Location = new System.Drawing.Point(30, 50);
            this.BR.Name = "BR";
            this.BR.Size = new System.Drawing.Size(46, 22);
            this.BR.TabIndex = 23;
            this.BR.Text = "BR";
            this.BR.UseVisualStyleBackColor = false;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.MaxRelativeTransmit);
            this.groupBox3.Controls.Add(this.InbandSpuriousEmissions);
            this.groupBox3.Controls.Add(this.DifferentialPhaseEncoding);
            this.groupBox3.Controls.Add(this.EDRCarrierFrequency);
            this.groupBox3.Controls.Add(this.MinRelativeTransmit);
            this.groupBox3.Location = new System.Drawing.Point(415, 319);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(403, 247);
            this.groupBox3.TabIndex = 22;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Test Item";
            // 
            // MaxRelativeTransmit
            // 
            this.MaxRelativeTransmit.AutoSize = true;
            this.MaxRelativeTransmit.Checked = true;
            this.MaxRelativeTransmit.CheckState = System.Windows.Forms.CheckState.Checked;
            this.MaxRelativeTransmit.Location = new System.Drawing.Point(29, 80);
            this.MaxRelativeTransmit.Name = "MaxRelativeTransmit";
            this.MaxRelativeTransmit.Size = new System.Drawing.Size(374, 22);
            this.MaxRelativeTransmit.TabIndex = 9;
            this.MaxRelativeTransmit.Text = "TRM/CA/10/C-EDR Relative MAX Transmit Power";
            this.MaxRelativeTransmit.UseVisualStyleBackColor = true;
            // 
            // InbandSpuriousEmissions
            // 
            this.InbandSpuriousEmissions.AutoSize = true;
            this.InbandSpuriousEmissions.Checked = true;
            this.InbandSpuriousEmissions.CheckState = System.Windows.Forms.CheckState.Checked;
            this.InbandSpuriousEmissions.Location = new System.Drawing.Point(29, 203);
            this.InbandSpuriousEmissions.Name = "InbandSpuriousEmissions";
            this.InbandSpuriousEmissions.Size = new System.Drawing.Size(366, 22);
            this.InbandSpuriousEmissions.TabIndex = 8;
            this.InbandSpuriousEmissions.Text = "TRM/CA/13/C-EDR In-band Spurious Emissions";
            this.InbandSpuriousEmissions.UseVisualStyleBackColor = true;
            // 
            // DifferentialPhaseEncoding
            // 
            this.DifferentialPhaseEncoding.AutoSize = true;
            this.DifferentialPhaseEncoding.Checked = true;
            this.DifferentialPhaseEncoding.CheckState = System.Windows.Forms.CheckState.Checked;
            this.DifferentialPhaseEncoding.Location = new System.Drawing.Point(29, 165);
            this.DifferentialPhaseEncoding.Name = "DifferentialPhaseEncoding";
            this.DifferentialPhaseEncoding.Size = new System.Drawing.Size(374, 22);
            this.DifferentialPhaseEncoding.TabIndex = 6;
            this.DifferentialPhaseEncoding.Text = "TRM/CA/12/C-EDR Differential Phase Encoding";
            this.DifferentialPhaseEncoding.UseVisualStyleBackColor = true;
            // 
            // EDRCarrierFrequency
            // 
            this.EDRCarrierFrequency.Checked = true;
            this.EDRCarrierFrequency.CheckState = System.Windows.Forms.CheckState.Checked;
            this.EDRCarrierFrequency.Location = new System.Drawing.Point(29, 108);
            this.EDRCarrierFrequency.Name = "EDRCarrierFrequency";
            this.EDRCarrierFrequency.Size = new System.Drawing.Size(365, 48);
            this.EDRCarrierFrequency.TabIndex = 4;
            this.EDRCarrierFrequency.Text = "TRM/CA/11/C EDR Carrier Frequency Stability and Modulation Accuracy";
            this.EDRCarrierFrequency.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.EDRCarrierFrequency.UseVisualStyleBackColor = true;
            // 
            // MinRelativeTransmit
            // 
            this.MinRelativeTransmit.AutoSize = true;
            this.MinRelativeTransmit.Checked = true;
            this.MinRelativeTransmit.CheckState = System.Windows.Forms.CheckState.Checked;
            this.MinRelativeTransmit.Location = new System.Drawing.Point(29, 38);
            this.MinRelativeTransmit.Name = "MinRelativeTransmit";
            this.MinRelativeTransmit.Size = new System.Drawing.Size(374, 22);
            this.MinRelativeTransmit.TabIndex = 2;
            this.MinRelativeTransmit.Text = "TRM/CA/10/C-EDR Relative MIN Transmit Power";
            this.MinRelativeTransmit.UseVisualStyleBackColor = true;
            // 
            // EDRTransmitter
            // 
            this.EDRTransmitter.Controls.Add(this.label12);
            this.EDRTransmitter.Controls.Add(this.T3DH5);
            this.EDRTransmitter.Controls.Add(this.T2DH1);
            this.EDRTransmitter.Controls.Add(this.T3DH3);
            this.EDRTransmitter.Controls.Add(this.T2DH3);
            this.EDRTransmitter.Controls.Add(this.T3DH1);
            this.EDRTransmitter.Controls.Add(this.T2DH5);
            this.EDRTransmitter.Location = new System.Drawing.Point(415, 175);
            this.EDRTransmitter.Name = "EDRTransmitter";
            this.EDRTransmitter.Size = new System.Drawing.Size(403, 125);
            this.EDRTransmitter.TabIndex = 21;
            this.EDRTransmitter.TabStop = false;
            this.EDRTransmitter.Text = "EDR Transmitter Test";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.BackColor = System.Drawing.Color.Orange;
            this.label12.Location = new System.Drawing.Point(26, 21);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(368, 18);
            this.label12.TabIndex = 20;
            this.label12.Text = "TX Test suite structure for Bluetooth 2.0+EDR";
            // 
            // T3DH5
            // 
            this.T3DH5.AutoSize = true;
            this.T3DH5.Checked = true;
            this.T3DH5.CheckState = System.Windows.Forms.CheckState.Checked;
            this.T3DH5.Location = new System.Drawing.Point(219, 90);
            this.T3DH5.Name = "T3DH5";
            this.T3DH5.Size = new System.Drawing.Size(70, 22);
            this.T3DH5.TabIndex = 19;
            this.T3DH5.Text = "3-DH5";
            this.T3DH5.UseVisualStyleBackColor = true;
            // 
            // T2DH1
            // 
            this.T2DH1.AutoSize = true;
            this.T2DH1.Checked = true;
            this.T2DH1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.T2DH1.Location = new System.Drawing.Point(29, 57);
            this.T2DH1.Name = "T2DH1";
            this.T2DH1.Size = new System.Drawing.Size(70, 22);
            this.T2DH1.TabIndex = 14;
            this.T2DH1.Text = "2-DH1";
            this.T2DH1.UseVisualStyleBackColor = true;
            // 
            // T3DH3
            // 
            this.T3DH3.AutoSize = true;
            this.T3DH3.Checked = true;
            this.T3DH3.CheckState = System.Windows.Forms.CheckState.Checked;
            this.T3DH3.Location = new System.Drawing.Point(124, 88);
            this.T3DH3.Name = "T3DH3";
            this.T3DH3.Size = new System.Drawing.Size(70, 22);
            this.T3DH3.TabIndex = 18;
            this.T3DH3.Text = "3-DH3";
            this.T3DH3.UseVisualStyleBackColor = true;
            // 
            // T2DH3
            // 
            this.T2DH3.AutoSize = true;
            this.T2DH3.Checked = true;
            this.T2DH3.CheckState = System.Windows.Forms.CheckState.Checked;
            this.T2DH3.Location = new System.Drawing.Point(124, 57);
            this.T2DH3.Name = "T2DH3";
            this.T2DH3.Size = new System.Drawing.Size(70, 22);
            this.T2DH3.TabIndex = 15;
            this.T2DH3.Text = "2-DH3";
            this.T2DH3.UseVisualStyleBackColor = true;
            // 
            // T3DH1
            // 
            this.T3DH1.AutoSize = true;
            this.T3DH1.Checked = true;
            this.T3DH1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.T3DH1.Location = new System.Drawing.Point(29, 86);
            this.T3DH1.Name = "T3DH1";
            this.T3DH1.Size = new System.Drawing.Size(70, 22);
            this.T3DH1.TabIndex = 17;
            this.T3DH1.Text = "3-DH1";
            this.T3DH1.UseVisualStyleBackColor = true;
            // 
            // T2DH5
            // 
            this.T2DH5.AutoSize = true;
            this.T2DH5.Checked = true;
            this.T2DH5.CheckState = System.Windows.Forms.CheckState.Checked;
            this.T2DH5.Location = new System.Drawing.Point(219, 57);
            this.T2DH5.Name = "T2DH5";
            this.T2DH5.Size = new System.Drawing.Size(70, 22);
            this.T2DH5.TabIndex = 16;
            this.T2DH5.Text = "2-DH5";
            this.T2DH5.UseVisualStyleBackColor = true;
            // 
            // TransmitterTest
            // 
            this.TransmitterTest.Controls.Add(this.label4);
            this.TransmitterTest.Controls.Add(this.T1DH3);
            this.TransmitterTest.Controls.Add(this.T1DH5);
            this.TransmitterTest.Controls.Add(this.T1DH1);
            this.TransmitterTest.Location = new System.Drawing.Point(415, 50);
            this.TransmitterTest.Name = "TransmitterTest";
            this.TransmitterTest.Size = new System.Drawing.Size(403, 100);
            this.TransmitterTest.TabIndex = 13;
            this.TransmitterTest.TabStop = false;
            this.TransmitterTest.Text = "Transmitter Test";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Orange;
            this.label4.Location = new System.Drawing.Point(23, 21);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(336, 18);
            this.label4.TabIndex = 0;
            this.label4.Text = "Tx Test suite structure for Bluetooth 1.2";
            // 
            // T1DH3
            // 
            this.T1DH3.AutoSize = true;
            this.T1DH3.Checked = true;
            this.T1DH3.CheckState = System.Windows.Forms.CheckState.Checked;
            this.T1DH3.Location = new System.Drawing.Point(106, 57);
            this.T1DH3.Name = "T1DH3";
            this.T1DH3.Size = new System.Drawing.Size(54, 22);
            this.T1DH3.TabIndex = 1;
            this.T1DH3.Text = "DH3";
            this.T1DH3.UseVisualStyleBackColor = true;
            // 
            // T1DH5
            // 
            this.T1DH5.AutoSize = true;
            this.T1DH5.Checked = true;
            this.T1DH5.CheckState = System.Windows.Forms.CheckState.Checked;
            this.T1DH5.Location = new System.Drawing.Point(199, 57);
            this.T1DH5.Name = "T1DH5";
            this.T1DH5.Size = new System.Drawing.Size(54, 22);
            this.T1DH5.TabIndex = 2;
            this.T1DH5.Text = "DH5";
            this.T1DH5.UseVisualStyleBackColor = true;
            // 
            // T1DH1
            // 
            this.T1DH1.AutoSize = true;
            this.T1DH1.Checked = true;
            this.T1DH1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.T1DH1.Location = new System.Drawing.Point(26, 57);
            this.T1DH1.Name = "T1DH1";
            this.T1DH1.Size = new System.Drawing.Size(54, 22);
            this.T1DH1.TabIndex = 0;
            this.T1DH1.Text = "DH1";
            this.T1DH1.UseVisualStyleBackColor = true;
            // 
            // TransmitterSetting
            // 
            this.TransmitterSetting.Controls.Add(this.TRMcarrierFrequency);
            this.TransmitterSetting.Controls.Add(this.TRMfrequenyRange);
            this.TransmitterSetting.Controls.Add(this.TRM20dBBandwidth);
            this.TransmitterSetting.Controls.Add(this.TRMinitialCarrier);
            this.TransmitterSetting.Controls.Add(this.TRMpowerControl);
            this.TransmitterSetting.Controls.Add(this.TRMadjacentChannel);
            this.TransmitterSetting.Controls.Add(this.TRMoutputPower);
            this.TransmitterSetting.Controls.Add(this.TRMpowerDensity);
            this.TransmitterSetting.Controls.Add(this.TRMmodulationCharacteristics);
            this.TransmitterSetting.Location = new System.Drawing.Point(30, 159);
            this.TransmitterSetting.Name = "TransmitterSetting";
            this.TransmitterSetting.Size = new System.Drawing.Size(364, 407);
            this.TransmitterSetting.TabIndex = 12;
            this.TransmitterSetting.TabStop = false;
            this.TransmitterSetting.Text = "TransmitterSetting";
            // 
            // TRMcarrierFrequency
            // 
            this.TRMcarrierFrequency.AutoSize = true;
            this.TRMcarrierFrequency.Checked = true;
            this.TRMcarrierFrequency.CheckState = System.Windows.Forms.CheckState.Checked;
            this.TRMcarrierFrequency.Location = new System.Drawing.Point(26, 366);
            this.TRMcarrierFrequency.Name = "TRMcarrierFrequency";
            this.TRMcarrierFrequency.Size = new System.Drawing.Size(310, 22);
            this.TRMcarrierFrequency.TabIndex = 9;
            this.TRMcarrierFrequency.Text = "TRM/CA/09/C-Carrier Frequency Drift";
            this.TRMcarrierFrequency.UseVisualStyleBackColor = true;
            // 
            // TRMfrequenyRange
            // 
            this.TRMfrequenyRange.AutoSize = true;
            this.TRMfrequenyRange.Checked = true;
            this.TRMfrequenyRange.CheckState = System.Windows.Forms.CheckState.Checked;
            this.TRMfrequenyRange.Location = new System.Drawing.Point(26, 162);
            this.TRMfrequenyRange.Name = "TRMfrequenyRange";
            this.TRMfrequenyRange.Size = new System.Drawing.Size(246, 22);
            this.TRMfrequenyRange.TabIndex = 4;
            this.TRMfrequenyRange.Text = "TRM/CA/04/C-Frequency Range";
            this.TRMfrequenyRange.UseVisualStyleBackColor = true;
            // 
            // TRM20dBBandwidth
            // 
            this.TRM20dBBandwidth.AutoSize = true;
            this.TRM20dBBandwidth.Checked = true;
            this.TRM20dBBandwidth.CheckState = System.Windows.Forms.CheckState.Checked;
            this.TRM20dBBandwidth.Location = new System.Drawing.Point(26, 204);
            this.TRM20dBBandwidth.Name = "TRM20dBBandwidth";
            this.TRM20dBBandwidth.Size = new System.Drawing.Size(238, 22);
            this.TRM20dBBandwidth.TabIndex = 5;
            this.TRM20dBBandwidth.Text = "TRM/CA/05/C-20dB Bandwidth";
            this.TRM20dBBandwidth.UseVisualStyleBackColor = true;
            // 
            // TRMinitialCarrier
            // 
            this.TRMinitialCarrier.AutoSize = true;
            this.TRMinitialCarrier.Checked = true;
            this.TRMinitialCarrier.CheckState = System.Windows.Forms.CheckState.Checked;
            this.TRMinitialCarrier.Location = new System.Drawing.Point(26, 326);
            this.TRMinitialCarrier.Name = "TRMinitialCarrier";
            this.TRMinitialCarrier.Size = new System.Drawing.Size(326, 22);
            this.TRMinitialCarrier.TabIndex = 8;
            this.TRMinitialCarrier.Text = "TRM/CA/08/C-Initial Carrier Frequency";
            this.TRMinitialCarrier.UseVisualStyleBackColor = true;
            // 
            // TRMpowerControl
            // 
            this.TRMpowerControl.AutoSize = true;
            this.TRMpowerControl.Checked = true;
            this.TRMpowerControl.CheckState = System.Windows.Forms.CheckState.Checked;
            this.TRMpowerControl.Location = new System.Drawing.Point(26, 119);
            this.TRMpowerControl.Name = "TRMpowerControl";
            this.TRMpowerControl.Size = new System.Drawing.Size(230, 22);
            this.TRMpowerControl.TabIndex = 3;
            this.TRMpowerControl.Text = "TRM/CA/03/C-Power control";
            this.TRMpowerControl.UseVisualStyleBackColor = true;
            // 
            // TRMadjacentChannel
            // 
            this.TRMadjacentChannel.AutoSize = true;
            this.TRMadjacentChannel.Checked = true;
            this.TRMadjacentChannel.CheckState = System.Windows.Forms.CheckState.Checked;
            this.TRMadjacentChannel.Location = new System.Drawing.Point(26, 242);
            this.TRMadjacentChannel.Name = "TRMadjacentChannel";
            this.TRMadjacentChannel.Size = new System.Drawing.Size(302, 22);
            this.TRMadjacentChannel.TabIndex = 6;
            this.TRMadjacentChannel.Text = "TRM/CA/06/C-Adjacent Channel Power";
            this.TRMadjacentChannel.UseVisualStyleBackColor = true;
            // 
            // TRMoutputPower
            // 
            this.TRMoutputPower.AutoSize = true;
            this.TRMoutputPower.BackColor = System.Drawing.Color.DarkGray;
            this.TRMoutputPower.Checked = true;
            this.TRMoutputPower.CheckState = System.Windows.Forms.CheckState.Checked;
            this.TRMoutputPower.Location = new System.Drawing.Point(26, 39);
            this.TRMoutputPower.Name = "TRMoutputPower";
            this.TRMoutputPower.Size = new System.Drawing.Size(222, 22);
            this.TRMoutputPower.TabIndex = 1;
            this.TRMoutputPower.Text = "TRM/CA/01/C-Output Power";
            this.TRMoutputPower.UseVisualStyleBackColor = false;
            // 
            // TRMpowerDensity
            // 
            this.TRMpowerDensity.AutoSize = true;
            this.TRMpowerDensity.Checked = true;
            this.TRMpowerDensity.CheckState = System.Windows.Forms.CheckState.Checked;
            this.TRMpowerDensity.Location = new System.Drawing.Point(26, 77);
            this.TRMpowerDensity.Name = "TRMpowerDensity";
            this.TRMpowerDensity.Size = new System.Drawing.Size(230, 22);
            this.TRMpowerDensity.TabIndex = 2;
            this.TRMpowerDensity.Text = "TRM/CA/02/C-Power Density";
            this.TRMpowerDensity.UseVisualStyleBackColor = true;
            // 
            // TRMmodulationCharacteristics
            // 
            this.TRMmodulationCharacteristics.AutoSize = true;
            this.TRMmodulationCharacteristics.Checked = true;
            this.TRMmodulationCharacteristics.CheckState = System.Windows.Forms.CheckState.Checked;
            this.TRMmodulationCharacteristics.Location = new System.Drawing.Point(26, 283);
            this.TRMmodulationCharacteristics.Name = "TRMmodulationCharacteristics";
            this.TRMmodulationCharacteristics.Size = new System.Drawing.Size(334, 22);
            this.TRMmodulationCharacteristics.TabIndex = 7;
            this.TRMmodulationCharacteristics.Text = "TRM/CA/07/C-Modulation Characteristics";
            this.TRMmodulationCharacteristics.UseVisualStyleBackColor = true;
            // 
            // Sett
            // 
            this.Sett.Controls.Add(this.Transmitter);
            this.Sett.Controls.Add(this.Receiver);
            this.Sett.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Sett.ItemSize = new System.Drawing.Size(100, 23);
            this.Sett.Location = new System.Drawing.Point(16, 176);
            this.Sett.Name = "Sett";
            this.Sett.SelectedIndex = 0;
            this.Sett.Size = new System.Drawing.Size(847, 616);
            this.Sett.TabIndex = 0;
            // 
            // MaxTransmintPower
            // 
            this.MaxTransmintPower.AutoSize = true;
            this.MaxTransmintPower.Location = new System.Drawing.Point(7, 22);
            this.MaxTransmintPower.Name = "MaxTransmintPower";
            this.MaxTransmintPower.Size = new System.Drawing.Size(230, 22);
            this.MaxTransmintPower.TabIndex = 0;
            this.MaxTransmintPower.Text = "Max Transmit Output Power";
            this.MaxTransmintPower.UseVisualStyleBackColor = true;
            // 
            // MinTransmitPower
            // 
            this.MinTransmitPower.AutoSize = true;
            this.MinTransmitPower.Location = new System.Drawing.Point(7, 50);
            this.MinTransmitPower.Name = "MinTransmitPower";
            this.MinTransmitPower.Size = new System.Drawing.Size(230, 22);
            this.MinTransmitPower.TabIndex = 1;
            this.MinTransmitPower.Text = "Min Transmit Output Power";
            this.MinTransmitPower.UseVisualStyleBackColor = true;
            // 
            // StartTest
            // 
            this.StartTest.BackColor = System.Drawing.Color.MediumSpringGreen;
            this.StartTest.Font = new System.Drawing.Font("Consolas", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.StartTest.Location = new System.Drawing.Point(213, 116);
            this.StartTest.Name = "StartTest";
            this.StartTest.Size = new System.Drawing.Size(232, 48);
            this.StartTest.TabIndex = 10;
            this.StartTest.Text = "Start Test";
            this.StartTest.UseVisualStyleBackColor = false;
            this.StartTest.Click += new System.EventHandler(this.StartTest_Click);
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.Color.DarkOrange;
            this.label5.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label5.Location = new System.Drawing.Point(511, 74);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(299, 25);
            this.label5.TabIndex = 11;
            this.label5.Text = "CMW info";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label8
            // 
            this.label8.BackColor = System.Drawing.Color.DarkOrange;
            this.label8.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(511, 28);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(75, 25);
            this.label8.TabIndex = 12;
            this.label8.Text = "CMW Port";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // CMWinfo
            // 
            this.CMWinfo.BackColor = System.Drawing.Color.White;
            this.CMWinfo.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CMWinfo.Location = new System.Drawing.Point(514, 116);
            this.CMWinfo.Name = "CMWinfo";
            this.CMWinfo.Size = new System.Drawing.Size(296, 74);
            this.CMWinfo.TabIndex = 14;
            this.CMWinfo.Text = "No CMW";
            this.CMWinfo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Port
            // 
            this.Port.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Port.FormattingEnabled = true;
            this.Port.Items.AddRange(new object[] {
            "RF1C",
            "RF3C"});
            this.Port.Location = new System.Drawing.Point(629, 27);
            this.Port.Name = "Port";
            this.Port.Size = new System.Drawing.Size(121, 26);
            this.Port.TabIndex = 16;
            this.Port.Text = "RF1C";
            // 
            // NunCh
            // 
            this.NunCh.AutoSize = false;
            this.NunCh.LargeChange = 1;
            this.NunCh.Location = new System.Drawing.Point(16, 6);
            this.NunCh.Maximum = 2;
            this.NunCh.Name = "NunCh";
            this.NunCh.Orientation = System.Windows.Forms.Orientation.Vertical;
            this.NunCh.RightToLeftLayout = true;
            this.NunCh.Size = new System.Drawing.Size(37, 104);
            this.NunCh.TabIndex = 1;
            this.NunCh.Value = 1;
            this.NunCh.Scroll += new System.EventHandler(this.NunCh_Scroll);
            // 
            // Channellist
            // 
            this.Channellist.Enabled = false;
            this.Channellist.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Channellist.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.Channellist.Location = new System.Drawing.Point(16, 116);
            this.Channellist.Name = "Channellist";
            this.Channellist.Size = new System.Drawing.Size(169, 25);
            this.Channellist.TabIndex = 17;
            this.Channellist.Text = "0,39,78";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(59, 84);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(15, 15);
            this.label9.TabIndex = 18;
            this.label9.Text = "1";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(59, 50);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(15, 15);
            this.label11.TabIndex = 20;
            this.label11.Text = "3";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(59, 18);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(23, 15);
            this.label13.TabIndex = 21;
            this.label13.Text = "78";
            // 
            // EventLog
            // 
            this.EventLog.BackColor = System.Drawing.SystemColors.InfoText;
            this.EventLog.Font = new System.Drawing.Font("Consolas", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EventLog.ForeColor = System.Drawing.SystemColors.Info;
            this.EventLog.Location = new System.Drawing.Point(870, 18);
            this.EventLog.Multiline = true;
            this.EventLog.Name = "EventLog";
            this.EventLog.Size = new System.Drawing.Size(515, 377);
            this.EventLog.TabIndex = 22;
            // 
            // ResultLog
            // 
            this.ResultLog.BackColor = System.Drawing.Color.PaleGreen;
            this.ResultLog.Font = new System.Drawing.Font("Consolas", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ResultLog.Location = new System.Drawing.Point(870, 401);
            this.ResultLog.Multiline = true;
            this.ResultLog.Name = "ResultLog";
            this.ResultLog.Size = new System.Drawing.Size(515, 387);
            this.ResultLog.TabIndex = 23;
            // 
            // BlueTooth
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1397, 804);
            this.Controls.Add(this.ResultLog);
            this.Controls.Add(this.EventLog);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.Channellist);
            this.Controls.Add(this.Port);
            this.Controls.Add(this.CMWinfo);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.StartTest);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.Att);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.GPIB);
            this.Controls.Add(this.NunCh);
            this.Controls.Add(this.Sett);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "BlueTooth";
            this.Text = "BT";
            this.Receiver.ResumeLayout(false);
            this.BERItem.ResumeLayout(false);
            this.BERItem.PerformLayout();
            this.Transmitter.ResumeLayout(false);
            this.Transmitter.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.EDRTransmitter.ResumeLayout(false);
            this.EDRTransmitter.PerformLayout();
            this.TransmitterTest.ResumeLayout(false);
            this.TransmitterTest.PerformLayout();
            this.TransmitterSetting.ResumeLayout(false);
            this.TransmitterSetting.PerformLayout();
            this.Sett.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.NunCh)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox Att;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox GPIB;
        private System.Windows.Forms.TabPage Receiver;
        private System.Windows.Forms.TabPage Transmitter;
        private System.Windows.Forms.CheckBox TRMpowerDensity;
        private System.Windows.Forms.CheckBox TRMcarrierFrequency;
        private System.Windows.Forms.CheckBox TRMinitialCarrier;
        private System.Windows.Forms.CheckBox TRMmodulationCharacteristics;
        private System.Windows.Forms.CheckBox TRMadjacentChannel;
        private System.Windows.Forms.CheckBox TRM20dBBandwidth;
        private System.Windows.Forms.CheckBox TRMfrequenyRange;
        private System.Windows.Forms.CheckBox TRMpowerControl;
        private System.Windows.Forms.CheckBox TRMoutputPower;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TabControl Sett;
        private System.Windows.Forms.Splitter splitter1;
        private System.Windows.Forms.CheckBox RCVsensitivityMeasurements;
        private System.Windows.Forms.CheckBox EDRmaximunInputLevel;
        private System.Windows.Forms.GroupBox BERItem;
        private System.Windows.Forms.CheckBox EDRBERfloorPerformance;
        private System.Windows.Forms.CheckBox EDRsensitivity;
        private System.Windows.Forms.CheckBox T1DH5;
        private System.Windows.Forms.CheckBox T1DH3;
        private System.Windows.Forms.CheckBox T1DH1;
        private System.Windows.Forms.GroupBox TransmitterSetting;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.CheckBox MaxRelativeTransmit;
        private System.Windows.Forms.CheckBox InbandSpuriousEmissions;
        private System.Windows.Forms.CheckBox DifferentialPhaseEncoding;
        private System.Windows.Forms.CheckBox EDRCarrierFrequency;
        private System.Windows.Forms.CheckBox MinRelativeTransmit;
        private System.Windows.Forms.GroupBox EDRTransmitter;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.CheckBox T3DH5;
        private System.Windows.Forms.CheckBox T2DH1;
        private System.Windows.Forms.CheckBox T3DH3;
        private System.Windows.Forms.CheckBox T2DH3;
        private System.Windows.Forms.CheckBox T3DH1;
        private System.Windows.Forms.CheckBox T2DH5;
        private System.Windows.Forms.GroupBox TransmitterTest;
        private System.Windows.Forms.CheckBox MaxTransmintPower;
        private System.Windows.Forms.CheckBox MinTransmitPower;
        private System.Windows.Forms.Button StartTest;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label CMWinfo;
        private System.Windows.Forms.ComboBox Port;
        private System.Windows.Forms.TrackBar NunCh;
        private System.Windows.Forms.TextBox Channellist;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.CheckBox LE;
        private System.Windows.Forms.CheckBox EDR;
        private System.Windows.Forms.CheckBox BR;
        private System.Windows.Forms.TextBox EventLog;
        private System.Windows.Forms.TextBox ResultLog;
        private System.Windows.Forms.CheckBox RCVMaximumInputLevel;
    }
}

