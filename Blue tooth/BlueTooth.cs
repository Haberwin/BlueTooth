using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.ExcelEdit;

namespace Blue_tooth
{

    public partial class BlueTooth : Form
    {
        public int Vi, i;
        public int ErrorStatus;
        public bool BlueHW, BlueSWEDR, BlueSWLE;
        public StringBuilder Feedback = new StringBuilder("", 3000);
        public string Path = Application.StartupPath;
        public string BDadress;
        public string DeviceName;
        List<int> listchannel = new List<int>();
        List<int> listACP = new List<int> { 3, 39, 75 };
        List<string> listBurtype = new List<string>();
        List<string> listBRtype = new List<string>();
        List<string> listEDRtype2 = new List<string>();
        List<string> listEDRtype3 = new List<string>();
        ExcelEdit Report = new ExcelEdit();
        public int Failcount;
        public string[] s;
        public BlueTooth()
        {
            InitializeComponent();
        }
        public bool GetVi()
        {
            string GPIB_Address;
            Output("Check GPIB port...");
            BlueSWEDR = false;
            BlueSWLE = false;
            BlueHW = false;

            ErrorStatus = -1;
            GPIB_Address = GPIB.Text;
            Visa32.viOpenDefaultRM(out int defrm);
            ErrorStatus = Visa32.viOpen(defrm, "GPIB0::" + GPIB_Address + "::INSTR", 0, 1000, out Vi);
            if (ErrorStatus != 0)
            {
                CMWinfo.Text = "GPIB端口无法识别,请检查连接设置!";
                CMWinfo.BackColor = Color.Red;
                Output("Cannot connect to the GPIB port!");
                Output("End test!");
                return false;
            }

            Printf(Vi, "*IDN?; *OPC?\n");
            Scanf(Vi, "%t", Feedback);
            if (Feedback.ToString().Contains("CMW"))
            {
                if (Feedback.ToString().Contains("145679"))
                {
                    CMWinfo.Text = Feedback.ToString();
                    CMWinfo.BackColor = Color.Green;
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "SYSTem:BASE:OPTion:LIST? HWOPtion\n");
                    Scanf(Vi, "%t", Feedback);
                    s = Feedback.ToString().Split(',');
                    for (i = 0; i < s.Length; i++)
                    {
                        if (s[i] == "H550B")
                        {
                            BlueHW = true;
                        }
                    }
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "SYSTem:BASE:OPTion:LIST? SWOPtion\n");
                    Scanf(Vi, "%t", Feedback);
                    s = Feedback.ToString().Split(',');
                    for (i = 0; i < s.Length; i++)
                    {
                        if (s[i] == "KS610")
                        {
                            BlueSWEDR = true;
                        }
                        if (s[i] == "KS611")
                        {
                            BlueSWEDR = true;
                        }
                    }

                    if (!BlueHW)
                    {
                        CMWinfo.Text = "缺失H550B组件!";
                        CMWinfo.BackColor = Color.Red;
                        return false;
                    }
                    if (!BlueSWEDR)
                    {
                        CMWinfo.Text = "缺失KS610组件,无法测试BR/EDR!";
                        CMWinfo.BackColor = Color.Red;
                        return false;
                    }
                    //Printf(Vi, "*RST; *OPC?\n");
                    //Printf(Vi, "*CLS; *OPC?\n");
                    Output("Successfully established connection");
                    return true;

                }
                else
                {
                    CMWinfo.Text = "CMW500序列号不符合,请联系刘文华!";
                    CMWinfo.BackColor = Color.Red;
                    return false;
                }

            }
            else
            {
                CMWinfo.Text = "CMW500设备无法找到，请确认连接设置！";
                CMWinfo.BackColor = Color.Red;
                return false;
            }

        }
        public bool ConnectCMW()
        {
            Output("Connect the test Phone......");
            Printf(Vi, "*RST; *WAI\n");
            Printf(Vi, "*CLS; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:OPMode RFT;; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BTYPe BR; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:HWINterface1 NONE; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:BRATe DH1; *WAI\n");
            Printf(Vi, "\n");
            Printf(Vi, "ROUTe:BLUetooth:SIGN:SCENario:OTRX " + Port.Text + ",RX1," + Port.Text + ",TX1; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:EATTenuation:OUTPut " + Att.Text + "; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:EATTenuation:INPut " + Att.Text + "; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:HOPPing OFF; *WAI\n");
            //Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:ARANging OFF;; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:ENPower 13; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:LEVel -30; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:UMARgin 3; *WAI\n");
            //Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BDADdress:CMW #H123456123456; *WAI\n");
            //Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:EPCMode AUTO; *WAI\n");
            Printf(Vi, "SOURce:BLUetooth:SIGN:STATe ON; *WAI\n");
            do
            {
                Output("Turn on the signal...");
                Thread.Sleep(500);
                Feedback.Remove(0, Feedback.Length);
                Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe:ALL?\n");
                Thread.Sleep(500);
                Scanf(Vi, "%t", Feedback);
                s = Feedback.ToString().Split(',');
            } while (s[0] != "ON");
            Output("Turn on the signal success!");
            //Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:INQuiry:ILENgth 10; *WAI\n");
            //Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:INQuiry:NOResponses 10; *WAI\n");
            Printf(Vi, "CALL:BLUetooth:SIGN:CONNection:ACTion INQuire; *WAI \n");
            Failcount = 0;
            do
            {
                Output("INQuire the Phone....");
                Thread.Sleep(500);
                Feedback.Remove(0, Feedback.Length);
                //Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:INQuiry:PTARgets:CATalog?\n");
                Thread.Sleep(500);
                Scanf(Vi, "%t", Feedback);
                s = Feedback.ToString().Split(',');
                Failcount += 1;
            } while (s[0] == "0" && Failcount < 40);
            if (Failcount >= 40)
            {
                
                Output("INQuire the Phone Failure");
                Output("No test device found!");
                return false;
            }
            Output("INQuire the Phone success!");
            /*
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:SVTimeout 16000; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PAGing:PSRMode R1; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PAGing:PTARget 1; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PAGing:TOUT 16384; *WAI\n");
            Printf(Vi, "CALL:BLUetooth:SIGN:CONNection:ACTion CONNect; \n");
            */
            Printf(Vi, "CALL:BLUetooth:SIGN:CONNection:ACTion SINQuiry; *WAI \n");
            Thread.Sleep(1500);
            Printf(Vi, "CALL:BLUetooth:SIGN:CONNection:ACTion TMConnect; *WAI\n");

            Failcount = 0;
            do
            {
                //
                Output("Connect Test Mode...");
                Thread.Sleep(1500);
                Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                Feedback.Remove(0, Feedback.Length);
                Scanf(Vi, "%t", Feedback);
                Failcount += 1;
            } while (Feedback.ToString() != "TCON\n" && Failcount <= 40);
            if (Failcount > 40)
            {
                Output("Connect Test Mode Failure!");
                return false;
            }
            Output("Connect Test Mode success!");
            Failcount = 0;
            Feedback.Remove(0, Feedback.Length);
            Printf(Vi, "SENSe:BLUetooth:SIGN:EUT:INFormation:BDADdress?\n");
            Thread.Sleep(500);
            Scanf(Vi, "%t", Feedback);
            BDadress = Feedback.ToString();
            Output("BDaddress:"+BDadress);
            Feedback.Remove(0, Feedback.Length);
            Printf(Vi, "SENSe:BLUetooth:SIGN:EUT:INFormation:NAME?\n");
            Thread.Sleep(500);
            Scanf(Vi, "%t", Feedback);
            DeviceName = Feedback.ToString();
            Output("DeviceName" + DeviceName);
            Printf(Vi, "ROUTe:BLUetooth:MEAS:SCENario:CSPath 'Bluetooth Sig1'\n");
            Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:REPetition CONTinuous; *WAI\n");
            Output("Connect the test Phone success!");
            return true;
        }
        public void GetOutputPower()
        {
            int row;
            Excel.Worksheet Txsheet;
            Output("Start Output Power......");
            Txsheet = Report.AddSheet("TRM-CA-01-C");
            Report.SetCellValue(Txsheet, 1, 1, "TRM/CA/01/C Output Power");
            //Report.SetCellProperty(Txsheet, 1, 1, 1, 1, 12, true, 3, -4119);
            Report.SetCellValue(Txsheet, 2, 2, "Device Name:");
            Report.SetCellValue(Txsheet, 2, 3, DeviceName);
            Report.SetCellValue(Txsheet, 3, 2, "BD Adress:");
            Report.SetCellValue(Txsheet, 3, 3, BDadress);
            Report.SetCellValue(Txsheet, 4, 2, "Burst Type:");
            Report.SetCellValue(Txsheet, 4, 3, "Basic Rate");
            Report.SetCellValue(Txsheet, 5, 2, "Package Type:");
            Report.SetCellValue(Txsheet, 5, 3, listBRtype[listBRtype.Count - 1]);
            Report.SetCellValue(Txsheet, 6, 2, "Pattern Type:");
            Report.SetCellValue(Txsheet, 6, 3, "PRBS9");
            Report.SetCellProperty(Txsheet, 2, 2, 6, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 8, 2, "Avarage Power limit:");
            Report.SetCellValue(Txsheet, 8, 3, "<20dBm");
            Report.SetCellValue(Txsheet, 9, 2, "Peak Power limit:");
            Report.SetCellValue(Txsheet, 9, 3, "<23dBm");
            Report.SetCellProperty(Txsheet, 8, 2, 9, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 11, 2, "Channel");
            Report.SetCellValue(Txsheet, 11, 3, "Avarage Power(dBm)");
            Report.SetCellValue(Txsheet, 11, 4, "Peak Power(dBm)");
            Report.SetCellProperty(Txsheet, 11, 2, 11, 4, 10, true, 6, -4119);
            ////Report.SetAutoFit(Txsheet, 2, 2, 11, 3);
            row = 11;
            if (listBRtype.Count != 0)
            {
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BTYPe BR; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:BRATe PRBS9; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:BRATe " + listBRtype[listBRtype.Count - 1] + "; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion MAX; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:RESult:ALL OFF,OFF,OFF,OFF,OFF,OFF,OFF,ON,OFF,OFF,OFF,OFF,OFF,OFF; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:REPetition CONTinuous; *WAI\n");
                foreach (int varCH in listchannel)
                {
                    row += 1;
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                    Scanf(Vi, "%t", Feedback);
                    if (Feedback.ToString() != "TCON\n")
                    {
                        if (!GetVi())
                        {
                            Application.Exit();
                            return;
                        }
                        if (!ConnectCMW())
                        {
                            CMWinfo.Text = "找不到测试设备，请检查连接设置！";
                            CMWinfo.BackColor = Color.Red;
                            return;
                        }
                    }
                    Output("Channel:"+varCH);
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback " + varCH.ToString() + ",0; *WAI\n");
                    Report.SetCellValue(Txsheet, row, 2, varCH.ToString());
                    Printf(Vi, "INITiate:BLUetooth:MEAS:MEValuation; *WAI\n");
                    Thread.Sleep(3000);
                    Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:PVTime:BRATe:AVERage?\n");
                    Failcount = 0;
                    do
                    {
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                        Failcount += 1;
                    } while (Feedback.Length == 0 && Failcount <= 20);
                    s = Feedback.ToString().Split(',');
                    if (s.Length >= 3)
                    {
                        Output("AVG Power:" + s[2]);
                        Report.SetCellValue(Txsheet, row, 3, s[2],"0.0");
                    }
                    else
                    {
                        Report.SetCellValue(Txsheet, row, 3, "Null");
                    }


                    Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:PVTime:BRATe:MAXimum?\n");
                    Failcount = 0;
                    do
                    {
                        Failcount += 1;
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                    } while (Feedback.Length == 0 && Failcount <= 20);
                    s = Feedback.ToString().Split(',');
                    if (s.Length >= 4)
                    {
                        Output("MAX Power:" + s[3]);
                        Report.SetCellValue(Txsheet, row, 4, s[3],"0.0");
                    }
                    else
                    {
                        Report.SetCellValue(Txsheet, row, 4, "Null");
                    }
                }
                Output("Test Output Power success");
                Report.SetCellProperty(Txsheet, 12, 2, row, 4, 9, false, 4, 1);
                //Report.SetFomules(Txsheet, 12, 3, row, 4,"0.0");
            }
        }
        public void GetPowerControl()
        {
            int row, col;
            double LastPower;
            Output("Start Output Control......");
            Excel.Worksheet Txsheet;
            Txsheet = Report.AddSheet("TRM-CA-03-C");
            Report.SetCellValue(Txsheet, 1, 1, "TRM/CA/03/C Power Control");
            Report.SetCellProperty(Txsheet, 1, 1, 1, 1, 12, true, 3, -4119);
            Report.SetCellValue(Txsheet, 2, 2, "Device Name:");
            Report.SetCellValue(Txsheet, 2, 3, DeviceName);
            Report.SetCellValue(Txsheet, 3, 2, "BD Adress:");
            Report.SetCellValue(Txsheet, 3, 3, BDadress);
            Report.SetCellValue(Txsheet, 4, 2, "Burst Type:");
            Report.SetCellValue(Txsheet, 4, 3, "Basic Rate");
            Report.SetCellValue(Txsheet, 5, 2, "Package Type:");
            Report.SetCellValue(Txsheet, 5, 3, listBRtype[listBRtype.Count - 1]);
            Report.SetCellValue(Txsheet, 6, 2, "Pattern Type:");
            Report.SetCellValue(Txsheet, 6, 3, "PRBS9");
            Report.SetCellProperty(Txsheet, 2, 2, 6, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 8, 2, "Step size limit:");
            Report.SetCellValue(Txsheet, 8, 3, "2dB<step size<6dB");
            Report.SetCellValue(Txsheet, 9, 2, "Min Power limit:");
            Report.SetCellValue(Txsheet, 9, 3, "Min Power<4dB");
            Report.SetCellProperty(Txsheet, 8, 2, 9, 2, 10, true, 6, -4119);
            //Report.SetAutoFit(Txsheet, 2, 2, 11, 2);
            //Report.SetAutoFit(Txsheet, 11, 3, 11, 5);
            Report.SetCellValue(Txsheet, 11, 2, "Channel");
            Report.SetCellValue(Txsheet, 11, 3, "Max Power(dBm)");

            row = 11;
            col = 3;
            if (listBRtype.Count != 0)
            {
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BTYPe BR; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:BRATe PRBS9; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:RESult:ALL OFF,OFF,OFF,OFF,OFF,OFF,OFF,ON,OFF,OFF,OFF,OFF,OFF,OFF; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:REPetition CONTinuous; *WAI\n");
                foreach (int varCH in listchannel)
                {
                    Output("Channel:"+varCH);
                    col = 3;
                    row += 1;
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                    Thread.Sleep(1000);
                    Scanf(Vi, "%t", Feedback);
                    if (Feedback.ToString() != "TCON\n")
                    {
                        if (!GetVi()){return;}
                        if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                    }
                    Report.SetCellValue(Txsheet, row, 2, varCH.ToString());
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback " + varCH.ToString() + ",0; *WAI\n");
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:BRATe " + listBRtype[listBRtype.Count - 1] + "; *WAI\n");
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion MAX\n");
                    Printf(Vi, "INITiate:BLUetooth:MEAS:MEValuation; *WAI\n");
                    Thread.Sleep(3000);
                    Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:PVTime:BRATe:AVERage?\n");
                    Failcount = 0;
                    do
                    {
                        Failcount += 1;
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                    } while (Feedback.Length == 0 && Failcount <= 20);
                    try
                    {
                        s = Feedback.ToString().Split(',');
                        Report.SetCellValue(Txsheet, row, 3, s[2],"0.0");
                        Output("Power at max level:" + s[2]);
                        LastPower = Convert.ToDouble(s[2]);
                    }
                    catch
                    {
                        Report.SetCellValue(Txsheet, row, 3, "Null");
                        break;
                    }

                    do
                    {
                        col += 1;
                        Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion DOWN; *WAI\n");
                        Output("Power level decline...");
                        Printf(Vi, "INITiate:BLUetooth:MEAS:MEValuation; *WAI\n");
                        Thread.Sleep(1500);
                        Feedback.Remove(0, Feedback.Length);
                        Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:PVTime:BRATe:AVERage?\n");
                        Failcount = 0;
                        do
                        {
                            Failcount += 1;
                            Thread.Sleep(500);
                            Scanf(Vi, "%t", Feedback);
                        } while (Feedback.Length == 0 && Failcount <= 20);
                        s = Feedback.ToString().Split(',');
                        if (s[0] != "0")
                        { break; }
                        try
                        {
                            Report.SetCellValue(Txsheet, row, col, (LastPower - Convert.ToDouble(s[2])).ToString(),"0.0");
                            LastPower = Convert.ToDouble(s[2]);
                            Output("Current power:" + s[2]);
                            Output("Power down step:" + (LastPower - Convert.ToDouble(s[2])).ToString());

                        }
                        catch
                        {
                            Report.SetCellValue(Txsheet, row, col, "Null");
                            break;
                        }
                        Feedback.Remove(0, Feedback.Length);
                        Printf(Vi, "SENSe:BLUetooth:SIGN:EUT:PCONtrol:STATe?\n");
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                        s = Feedback.ToString().Split(',');
                        Report.SetCellValue(Txsheet, 11, col, s[1], "0.0");
                    } while (s[1] != "MIN");
                    col += 1;
                    Report.SetCellValue(Txsheet, 11, col, "Min Power");
                    Report.SetCellValue(Txsheet, row, col, LastPower.ToString(), "0.0");

                }
                Report.SetCellProperty(Txsheet, 11, 2, 11, col, 10, true, 6, -4119);
                Report.SetCellProperty(Txsheet, 12, 2, row, col, 9, false, 4, 1);
                //Report.SetFomules(Txsheet, 12, 3, row, col,"0.0");
            }

        }
        public void GetFrequanceRange()
        {
            int row;
            Excel.Worksheet Txsheet;
            Output("Start Frequency Range test......");
            Txsheet = Report.AddSheet("TRM-CA-04-C");
            Report.SetCellValue(Txsheet, 1, 1, "TRM/CA/04/C TX Output Spectrum-Frequency Range");
            //Report.SetCellProperty(Txsheet, 1, 1, 1, 1, 12, true, 3, -4119);
            Report.SetCellValue(Txsheet, 2, 2, "Device Name:");
            Report.SetCellValue(Txsheet, 2, 3, DeviceName);
            Report.SetCellValue(Txsheet, 3, 2, "BD Adress:");
            Report.SetCellValue(Txsheet, 3, 3, BDadress);
            Report.SetCellValue(Txsheet, 4, 2, "Burst Type:");
            Report.SetCellValue(Txsheet, 4, 3, "Basic Rate");
            Report.SetCellValue(Txsheet, 5, 2, "Package Type:");
            Report.SetCellValue(Txsheet, 5, 3, listBRtype[listBRtype.Count - 1]);
            Report.SetCellValue(Txsheet, 6, 2, "Pattern Type:");
            Report.SetCellValue(Txsheet, 6, 3, "PRBS9");
            Report.SetCellProperty(Txsheet, 2, 2, 6, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 8, 2, "Channel 0 limit:");
            Report.SetCellValue(Txsheet, 8, 3, ">2400MHz");
            Report.SetCellValue(Txsheet, 9, 2, "Channel 78 limit:");
            Report.SetCellValue(Txsheet, 9, 3, "<2483.5MHz");
            Report.SetCellProperty(Txsheet, 8, 2, 9, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 11, 2, "Channel");
            Report.SetCellValue(Txsheet, 11, 3, "FL(MHz)");
            Report.SetCellValue(Txsheet, 11, 4, "FH(MHz)");
            Report.SetCellProperty(Txsheet, 11, 2, 11, 4, 10, true, 6, -4119);
            //Report.SetAutoFit(Txsheet, 2, 2, 11, 2);
            //Report.SetAutoFit(Txsheet, 11, 3, 11, 5);
            row = 11;
            if (listBRtype.Count != 0)
            {
               
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BTYPe BR; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:BRATe PRBS9; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:BRATe " + listBRtype[listBRtype.Count - 1] + "; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:RESult:ALL OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,ON,OFF; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:REPetition CONTinuous; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion MAX; *WAI\n");

                row += 1;
                Feedback.Remove(0, Feedback.Length);
                Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                Thread.Sleep(500);
                Scanf(Vi, "%t", Feedback);
                if (Feedback.ToString() != "TCON\n")
                {
                    if (!GetVi()){return;}
                    if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                }
                Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback 0,0; *WAI\n");
                Output("Channel:0");
                Report.SetCellValue(Txsheet, row, 2, "0");
                Printf(Vi, "INITiate:BLUetooth:MEAS:MEValuation; *WAI\n");
                Thread.Sleep(3000);
                Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:FRANge:BRATe:CURRent?\n");
                Failcount = 0;
                do
                {
                    Failcount += 1;
                    Feedback.Remove(0, Feedback.Length);
                    Thread.Sleep(500);
                    Scanf(Vi, "%t", Feedback);
                } while (Feedback.Length == 0 && Failcount <= 20);
                s = Feedback.ToString().Split(',');
                try
                {
                    
                    Report.SetCellValue(Txsheet, row, 3, Convert.ToDouble(s[3])/1000000.0, "0.0");
                    Report.SetCellValue(Txsheet, row, 4, Convert.ToDouble(s[4]) / 1000000.0, "0.0");
                    Output("FL"+ (Convert.ToDouble(s[3]) / 1000000.0).ToString());
                    Output("FH" + (Convert.ToDouble(s[4]) / 1000000.0).ToString());
                }
                catch
                {
                    Report.SetCellValue(Txsheet, row, 3, "Null");
                    Report.SetCellValue(Txsheet, row, 4, "Null");
                }
                row += 1;
                Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback 78,0; *WAI\n");
                Output("Channel:78");
                Report.SetCellValue(Txsheet, row, 2, "78");
                Printf(Vi, "INITiate:BLUetooth:MEAS:MEValuation; *WAI\n");
                Thread.Sleep(1500);
                Feedback.Remove(0, Feedback.Length);
                Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:FRANge:BRATe:CURRent?\n");
                Failcount = 0;
                do
                {
                    Failcount += 1;
                    Thread.Sleep(500);
                    Scanf(Vi, "%t", Feedback);
                } while (Feedback.Length == 0 && Failcount <= 20);
                s = Feedback.ToString().Split(',');
                if (s.Length >= 4)
                {

                    Report.SetCellValue(Txsheet, row, 3, Convert.ToDouble(s[3]) / 1000000.0, "0.0");
                    Report.SetCellValue(Txsheet, row, 4, Convert.ToDouble(s[4]) / 1000000.0, "0.0");
                    Output("FL" + (Convert.ToDouble(s[3]) / 1000000.0).ToString());
                    Output("FH" + (Convert.ToDouble(s[4]) / 1000000.0).ToString());
                }
                else
                {
                    Report.SetCellValue(Txsheet, row, 3, "Null");
                    Report.SetCellValue(Txsheet, row, 4, "Null");
                }
                Report.SetCellProperty(Txsheet, 12, 2, row, 4, 9, false, 4, 1);
                //Report.SetFomules(Txsheet, 12, 3, row, 4, "0.0");
            }

        }
        public void Get20dBBandwidth()
        {
            int row;
            Excel.Worksheet Txsheet;
            Output("Start 20dB Bandwidth ......");
            Txsheet = Report.AddSheet("TRM-CA-05-C");
            Report.SetCellValue(Txsheet, 1, 1, "TRM/CA/05/C TX Output Spectrum-20dB Bandwidth");
            //Report.SetCellProperty(Txsheet, 1, 1, 1, 1, 12, true, 3, -4119);
            Report.SetCellValue(Txsheet, 2, 2, "Device Name:");
            Report.SetCellValue(Txsheet, 2, 3, DeviceName);
            Report.SetCellValue(Txsheet, 3, 2, "BD Adress:");
            Report.SetCellValue(Txsheet, 3, 3, BDadress);
            Report.SetCellValue(Txsheet, 4, 2, "Burst Type:");
            Report.SetCellValue(Txsheet, 4, 3, "Basic Rate");
            Report.SetCellValue(Txsheet, 5, 2, "Package Type:");
            Report.SetCellValue(Txsheet, 5, 3, listBRtype[listBRtype.Count - 1]);
            Report.SetCellValue(Txsheet, 6, 2, "Pattern Type:");
            Report.SetCellValue(Txsheet, 6, 3, "PRBS9");
            Report.SetCellProperty(Txsheet, 2, 2, 6, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 8, 2, "If PeakPower>=0dBm:");
            Report.SetCellValue(Txsheet, 8, 3, "fH-fL<=1.5MHz");
            Report.SetCellValue(Txsheet, 9, 2, "If PeakPower<0dBm:");
            Report.SetCellValue(Txsheet, 9, 3, "fH-fL<=1MHz");
            Report.SetCellProperty(Txsheet, 8, 2, 9, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 11, 2, "Channel");
            Report.SetCellValue(Txsheet, 11, 3, "fL(MHz)");
            Report.SetCellValue(Txsheet, 11, 4, "fH(MHz)");
            Report.SetCellValue(Txsheet, 11, 5, "fL-fH(MHz)");
            //Report.SetCellValue(Txsheet, 11, 4, "Peak Power(dBm)");
            Report.SetCellProperty(Txsheet, 11, 2, 11, 5, 10, true, 6, -4119);
            //Report.SetAutoFit(Txsheet, 2, 2, 11, 2);
            //Report.SetAutoFit(Txsheet, 11, 3, 11, 5);
            row = 11;
            if (listBRtype.Count != 0)
            {
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BTYPe BR; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:BRATe PRBS9; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:BRATe " + listBRtype[listBRtype.Count - 1] + "; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:RESult:ALL OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,ON,OFF,OFF,OFF,OFF; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:REPetition CONTinuous; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion MAX; *WAI\n");
                foreach (int varCH in listchannel)
                {
                    Output("Channel:" + varCH.ToString());
                    row += 1;
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                    Thread.Sleep(500);
                    Scanf(Vi, "%t", Feedback);
                    if (Feedback.ToString() != "TCON\n")
                    {
                        if (!GetVi()){return;}
                        if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                    }
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback " + varCH.ToString() + ",0; *WAI\n");
                    Report.SetCellValue(Txsheet, row, 2, varCH.ToString());

                    Printf(Vi, "INITiate:BLUetooth:MEAS:MEValuation; *WAI\n");
                    Thread.Sleep(3000);

                    Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:SOBW:BRATe:MAXimum?\n");
                    Failcount = 0;
                    do
                    {
                        Failcount += 1;
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                    } while (Feedback.Length == 0 && Failcount <= 20);
                    s = Feedback.ToString().Split(',');
                    try
                    {
                        Report.SetCellValue(Txsheet, row, 3, Convert.ToDouble(s[4]) / 1000000.0,"0.00");
                        Report.SetCellValue(Txsheet, row, 4, Convert.ToDouble(s[5]) / 1000000.0,"0.00");
                        Report.SetCellValue(Txsheet, row, 5, Convert.ToDouble(s[6]) / 1000000.0,"0.00");
                        Output("FL:"+(Convert.ToDouble(s[4]) / 1000000.0).ToString());
                        Output("FH:" + (Convert.ToDouble(s[5]) / 1000000.0).ToString());
                        Output("FH-FL:" + (Convert.ToDouble(s[6]) / 1000000.0).ToString());

                    }
                    catch
                    {
                        Report.SetCellValue(Txsheet, row, 3, "Null");
                        Report.SetCellValue(Txsheet, row, 4, "Null");
                        Report.SetCellValue(Txsheet, row, 5, "Null");
                    }

                }
                Report.SetCellProperty(Txsheet, 12, 2, row, 5, 9, false, 4, 1);
                //Report.SetFomules(Txsheet, 12, 3, row, 5,"0.00");
            }

        }
        public void GetAdjacentChannelPower()
        {
            int row, col, TempI;
            Excel.Worksheet Txsheet;
            Output("Start Adjacent Channel Power......");
            Txsheet = Report.AddSheet("TRM-CA-06-C");
            Report.SetCellValue(Txsheet, 1, 1, "TRM/CA/06/C TX Output Spectrum Adjacent Channel Power");
            //Report.SetCellProperty(Txsheet, 1, 1, 1, 1, 12, true, 3, -4119);
            Report.SetCellValue(Txsheet, 2, 2, "Device Name:");
            Report.SetCellValue(Txsheet, 2, 3, DeviceName);
            Report.SetCellValue(Txsheet, 3, 2, "BD Adress:");
            Report.SetCellValue(Txsheet, 3, 3, BDadress);
            Report.SetCellValue(Txsheet, 4, 2, "Burst Type:");
            Report.SetCellValue(Txsheet, 4, 3, "Basic Rate");
            Report.SetCellValue(Txsheet, 5, 2, "Package Type:");
            Report.SetCellValue(Txsheet, 5, 3, "DH1");
            Report.SetCellValue(Txsheet, 6, 2, "Pattern Type:");
            Report.SetCellValue(Txsheet, 6, 3, "PRBS9");
            Report.SetCellProperty(Txsheet, 2, 2, 6, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 8, 2, "If |M-N|=2:");
            Report.SetCellValue(Txsheet, 8, 3, "Limit<-20dBm");
            Report.SetCellValue(Txsheet, 9, 2, "If |M-N|>=3:");
            Report.SetCellValue(Txsheet, 9, 3, "Limit<-40dBm");
            Report.SetCellProperty(Txsheet, 8, 2, 9, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 11, 2, "Channel");
            Report.SetCellValue(Txsheet, 11, 3, "CH3");
            Report.SetCellValue(Txsheet, 11, 4, "CH39");
            Report.SetCellValue(Txsheet, 11, 5, "CH75");
            Report.SetCellProperty(Txsheet, 11, 2, 11, 5, 10, true, 6, -4119);
            //Report.SetAutoFit(Txsheet, 2, 2, 11, 2);
            //Report.SetAutoFit(Txsheet, 11, 3, 11, 5);
            row = 11;
            col = 2;
            if (listBRtype.Count != 0)
            {
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BTYPe BR; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:BRATe DH1; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:BRATe PRBS9; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:RESult:ALL OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,ON,OFF,OFF,OFF; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion MAX; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:REPetition CONTinuous; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:SACP:BRATe:MEASurement:MODE CH79; *WAI\n");
                foreach (int varCH in listACP)
                {
                    Output("Channel:" + varCH.ToString());
                    col += 1;
                    row = 12;
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                    Scanf(Vi, "%t", Feedback);
                    if (Feedback.ToString() != "TCON\n")
                    {
                        if (!GetVi()){return;}
                        if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                    }
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback " + varCH.ToString() + ",0; *WAI\n");
                    //Report.SetCellValue(Txsheet, row, 2, varCH.ToString());
                    Printf(Vi, "INITiate:BLUetooth:MEAS:MEValuation; *WAI\n");

                    Failcount = 0;
                    do
                    {
                        Failcount += 1;
                        Feedback.Remove(0, Feedback.Length);
                        Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:STATe?\n");
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                        s = Feedback.ToString().Split(',');
                    } while (!(s[0] == "RUN" || Failcount <= 20));
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "*OPC?\n");
                    Thread.Sleep(500);
                    Scanf(Vi, "%t", Feedback);

                    Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:SACP:BRATe?\n");
                    Failcount = 0;
                    do
                    {
                        Failcount += 1;
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                    } while (Feedback.Length == 0 && Failcount <= 20);
                    s = Feedback.ToString().Split(',');
                    Report.SetCellValue(Txsheet, row, 2, "Number of exceptions:");
                    if (s.Length >= 82)
                    {
                        Report.SetCellValue(Txsheet, row, col, s[3]);
                    }
                    else
                    {
                        Report.SetCellValue(Txsheet, row, col, "Null");
                        break;
                    }

                    row = 13;
                    for (TempI = 0; TempI < 79; TempI++, row++)
                    {
                        Report.SetCellValue(Txsheet, row, 2, TempI.ToString());
                        Report.SetCellValue(Txsheet, row, col, s[TempI + 4],"0.00");
                        Output("CH" + TempI.ToString() + "\t" + s[TempI + 4]);
                    }
                }
                Report.SetCellProperty(Txsheet, 12, 2, row - 1, 5, 9, false, 4, 1);
                //Report.SetFomules(Txsheet, 12, 3, row, 5,"0.00");
            }
            Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:SACP:BRATe:MEASurement:MODE CH21; *WAI\n");
        }
        public void GetModulationCharacterristics()
        {
            int row;
            string Tempf1avg;
            Output("Start Modulation Characterristics .......");
            Excel.Worksheet Txsheet;
            Txsheet = Report.AddSheet("TRM-CA-07-C");
            Report.SetCellValue(Txsheet, 1, 1, "TRM/CA/07/C ModulationCharacterristics");
            //Report.SetCellProperty(Txsheet, 1, 1, 1, 1, 12, true, 3, -4119);
            Report.SetCellValue(Txsheet, 2, 2, "Device Name:");
            Report.SetCellValue(Txsheet, 2, 3, DeviceName);
            Report.SetCellValue(Txsheet, 3, 2, "BD Adress:");
            Report.SetCellValue(Txsheet, 3, 3, BDadress);
            Report.SetCellValue(Txsheet, 4, 2, "Burst Type:");
            Report.SetCellValue(Txsheet, 4, 3, "Basic Rate");
            Report.SetCellValue(Txsheet, 5, 2, "Package Type:");
            Report.SetCellValue(Txsheet, 5, 3, listBRtype[listBRtype.Count - 1]);
            Report.SetCellValue(Txsheet, 6, 2, "Pattern Type:");
            Report.SetCellValue(Txsheet, 6, 3, "11110000 for f1,10101010 for f2");
            Report.SetCellProperty(Txsheet, 2, 2, 6, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 8, 2, "Limit:");
            Report.SetCellValue(Txsheet, 8, 3, "140KHz<=f1_avg<=175KHz");
            Report.SetCellValue(Txsheet, 9, 2, "Limit:");
            Report.SetCellValue(Txsheet, 9, 3, "f2_max>=115KHz for at least 99.9% of all f2-max");
            Report.SetCellValue(Txsheet, 10, 2, "Limit:");
            Report.SetCellValue(Txsheet, 10, 3, "f2_avg/f1_avg >=0.8");
            Report.SetCellProperty(Txsheet, 8, 2, 10, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 12, 2, "Channel");
            Report.SetCellValue(Txsheet, 12, 3, "Feq Dev f1_avg(KHz)");
            Report.SetCellValue(Txsheet, 12, 4, "Feq Dev f2_avg(KHz)");
            Report.SetCellValue(Txsheet, 12, 5, "Feq Dev f2_max(KHz)");
            Report.SetCellValue(Txsheet, 12, 6, "f2_avg/f1_avg");
            Report.SetCellProperty(Txsheet, 12, 2, 12, 6, 10, true, 6, -4119);
            row = 12;
            if (listBRtype.Count != 0)
            {
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BTYPe BR; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:BRATe " + listBRtype[listBRtype.Count - 1] + "; *WAI\n");
                //Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:BRATe P44; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:RESult:ALL OFF,OFF,ON,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion MAX; *WAI\n");
                foreach (int varCH in listchannel)
                {
                    Output("Channel:" + varCH.ToString());
                    row += 1;
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                    Scanf(Vi, "%t", Feedback);
                    if (Feedback.ToString() != "TCON\n")
                    {
                        if (!GetVi()){return;}
                        if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                    }
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback " + varCH.ToString() + ",0; *WAI\n");
                    Report.SetCellValue(Txsheet, row, 2, varCH.ToString());
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:BRATe P44; *WAI\n");
                    Printf(Vi, "INITiate:BLUetooth:MEAS:MEValuation; *WAI\n");
                    Thread.Sleep(3000);

                    Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:MODulation:BRATe:AVERage?\n");
                    Failcount = 0;
                    do
                    {
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                        Failcount += 1;
                    } while (Feedback.Length == 0 && Failcount <= 20);
                    s = Feedback.ToString().Split(',');
                    Tempf1avg = "";
                    try
                    {
                        Report.SetCellValue(Txsheet, row, 3, Convert.ToDouble(s[6]) / 1000.0, "0.0");
                        Output("f1-avg:" + s[6]);
                        Tempf1avg = s[6];
                    }
                    catch
                    {
                        Report.SetCellValue(Txsheet, row, 3, "Null");
                    }
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:BRATe P11; *WAI\n");
                    Printf(Vi, "INITiate:BLUetooth:MEAS:MEValuation; *WAI\n");
                    Thread.Sleep(3000);
                    Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:MODulation:BRATe:AVERage?\n");
                    Failcount = 0;
                    do
                    {
                        Failcount += 1;
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                    } while (Feedback.Length == 0 && Failcount <= 20);
                    s = Feedback.ToString().Split(',');
                    try
                    {
                        Report.SetCellValue(Txsheet, row, 4, Convert.ToDouble(s[9]) / 1000.0, "0.0");
                        Report.SetCellValue(Txsheet, row, 5, Convert.ToDouble(s[11]) / 1000.0, "0.0");
                        Output("f2-avg:" + s[9]);
                        Output("f2-max:" + s[10]);
                        try
                        {
                            Report.SetCellValue(Txsheet, row, 6, (Convert.ToDouble(s[9]) / Convert.ToDouble(Tempf1avg)).ToString(), "0.00");
                            Output("f2-avg/f1-avg:" + (Convert.ToDouble(s[9]) / Convert.ToDouble(Tempf1avg)).ToString());
                        }
                        catch { Report.SetCellValue(Txsheet, row, 6, "Null"); }
                    }
                    catch
                    {
                        Report.SetCellValue(Txsheet, row, 4, "Null");
                        Report.SetCellValue(Txsheet, row, 5, "Null");

                    }
                }
                Report.SetCellProperty(Txsheet, 13, 2, row, 6, 9, false, 4, 1);
                //Report.SetFomules(Txsheet, 13, 3, row, 6,"0,00");
            }

        }
        public void GetCarrierFrequencyTolerance()
        {
            int row;
            Excel.Worksheet Txsheet;
            Output("Start Carrier Frequency Tolerance......");
            Txsheet = Report.AddSheet("TRM-CA-08-C");
            Report.SetCellValue(Txsheet, 1, 1, "TRM/CA/08/C Initial Carrier Frequency Tolerance");
            //Report.SetCellProperty(Txsheet, 1, 1, 1, 1, 12, true, 3, -4119);
            Report.SetCellValue(Txsheet, 2, 2, "Device Name:");
            Report.SetCellValue(Txsheet, 2, 3, DeviceName);
            Report.SetCellValue(Txsheet, 3, 2, "BD Adress:");
            Report.SetCellValue(Txsheet, 3, 3, BDadress);
            Report.SetCellValue(Txsheet, 4, 2, "Burst Type:");
            Report.SetCellValue(Txsheet, 4, 3, "Basic Rate");
            Report.SetCellValue(Txsheet, 5, 2, "Package Type:");
            Report.SetCellValue(Txsheet, 5, 3, "DH1");
            Report.SetCellValue(Txsheet, 6, 2, "Pattern Type:");
            Report.SetCellValue(Txsheet, 6, 3, "PRBS9");
            Report.SetCellProperty(Txsheet, 2, 2, 7, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 7, 2, "Hopping");
            Report.SetCellValue(Txsheet, 7, 3, "ON");
            Report.SetCellValue(Txsheet, 9, 2, "Limit:");
            Report.SetCellValue(Txsheet, 9, 3, "|Freq.accuracy|<=75KHz");
            Report.SetCellProperty(Txsheet, 9, 2, 9, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 11, 2, "Channel");
            Report.SetCellValue(Txsheet, 11, 3, "Freq.accuracy");
            Report.SetCellProperty(Txsheet, 11, 2, 11, 3, 10, true, 6, -4119);
            //Report.SetAutoFit(Txsheet, 2, 2, 11, 2);
            row = 11;
            if (listBRtype.Count != 0)
            {
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BTYPe BR; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:BRATe DH1; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:HOPPing ON; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:BRATe PRBS9; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:RESult:ALL OFF,OFF,ON,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion MAX; *WAI\n");
                foreach (int varCH in listchannel)
                {
                    Output("Channel:" + varCH.ToString());
                    row += 1;
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                    Scanf(Vi, "%t", Feedback);
                    if (Feedback.ToString() != "TCON\n")
                    {
                        if (!GetVi()){return;}
                        if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                    }
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback " + varCH.ToString() + ",0; *WAI\n");
                    Report.SetCellValue(Txsheet, row, 2, varCH.ToString());
                    Printf(Vi, "INITiate:BLUetooth:MEAS:MEValuation; *WAI\n");
                    Thread.Sleep(3000);

                    Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:MODulation:BRATe:AVERage?\n");
                    Failcount = 0;
                    do
                    {
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                        Failcount += 1;
                    } while (Feedback.Length == 0 && Failcount <= 20);
                    s = Feedback.ToString().Split(',');
                    try
                    {
                        Report.SetCellValue(Txsheet, row, 3, Convert.ToDouble(s[3]) / 1000.0, "0.0");
                        Output("Freq.accuracy:" + s[3]);
                    }
                    catch
                    {
                        Report.SetCellValue(Txsheet, row, 3, "Null");
                    }
                }
                Report.SetCellProperty(Txsheet, 12, 2, row, 3, 9, false, 4, 1);
                //Report.SetFomules(Txsheet, 12, 3, row, 3,"0.0");
            }

        }
        public void GetCarrierFrequancyDirft()
        {
            int row, col;
            Excel.Worksheet Txsheet;
            Output("Start Carrier Frequancy Dirft");
            Txsheet = Report.AddSheet("TRM-CA-09-C");
            Report.SetCellValue(Txsheet, 1, 1, "TRM/CA/09/C Carrier Frequency Dirft");
            //Report.SetCellProperty(Txsheet, 1, 1, 1, 1, 12, true, 3, -4119);
            Report.SetCellValue(Txsheet, 2, 2, "Device Name:");
            Report.SetCellValue(Txsheet, 2, 3, DeviceName);
            Report.SetCellValue(Txsheet, 3, 2, "BD Adress:");
            Report.SetCellValue(Txsheet, 3, 3, BDadress);
            Report.SetCellValue(Txsheet, 4, 2, "Burst Type:");
            Report.SetCellValue(Txsheet, 4, 3, "Basic Rate");
            Report.SetCellValue(Txsheet, 5, 2, "Package Type:");
            Report.SetCellValue(Txsheet, 5, 3, "DH1");
            Report.SetCellValue(Txsheet, 5, 5, "DH3");
            Report.SetCellValue(Txsheet, 5, 7, "DH5");
            Report.SetCellValue(Txsheet, 6, 2, "Pattern Type:");
            Report.SetCellValue(Txsheet, 6, 3, "101010");
            Report.SetCellProperty(Txsheet, 2, 2, 7, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 7, 2, "Hopping");
            Report.SetCellValue(Txsheet, 7, 3, "ON");
            Report.SetCellValue(Txsheet, 9, 2, "Limit:");
            Report.SetCellValue(Txsheet, 9, 3, "|Freq.Drity|<=25KHz");
            Report.SetCellValue(Txsheet, 9, 5, "|Freq.Drity|<=40KHz");
            Report.SetCellValue(Txsheet, 9, 7, "|Freq.Drity|<=40KHz");
            Report.SetCellProperty(Txsheet, 9, 2, 9, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 11, 2, "Channel");
            Report.SetCellValue(Txsheet, 11, 3, "Freq.Drity DH1");
            Report.SetCellValue(Txsheet, 11, 5, "Freq.Drity DH3");
            Report.SetCellValue(Txsheet, 11, 7, "Freq.Drity DH5");
            Report.SetCellProperty(Txsheet, 11, 2, 11, 3, 10, true, 6, -4119);
            Report.SetCellProperty(Txsheet, 11, 5, 11, 5, 10, true, 6, -4119);
            Report.SetCellProperty(Txsheet, 11, 7, 11, 7, 10, true, 6, -4119);
            //Report.SetAutoFit(Txsheet, 2,2, 11, 7);
            row = 11;
            col = 1;
            listBRtype.Clear();
            listBRtype.Add("DH1");
            listBRtype.Add("DH3");
            listBRtype.Add("DH5");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BTYPe BR; *WAI\n");

            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:HOPPing ON; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:BRATe P11; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:RESult:ALL OFF,OFF,ON,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion MAX; *WAI\n");
            foreach (string varBR in listBRtype)
            {
                Output("Packets type:" + varBR);
                col += 2;
                row = 11;
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:BRATe " + varBR + "; *WAI\n");
                foreach (int varCH in listchannel)
                {
                    Output("Channel:" + varCH.ToString());
                    row += 1;
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                    Scanf(Vi, "%t", Feedback);
                    if (Feedback.ToString() != "TCON\n")
                    {
                        if (!GetVi()){return;}
                        if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                    }
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback " + varCH.ToString() + ",0; *WAI\n");
                    Report.SetCellValue(Txsheet, row, 2, varCH.ToString());
                    Printf(Vi, "INITiate:BLUetooth:MEAS:MEValuation; *WAI\n");
                    Thread.Sleep(3000);
                    Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:MODulation:BRATe:AVERage?\n");
                    Failcount = 0;
                    do
                    {
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                        Failcount += 1;
                    } while (Feedback.Length == 0 && Failcount <= 20);
                    s = Feedback.ToString().Split(',');
                    try
                    {
                        Report.SetCellValue(Txsheet, row, col, Convert.ToDouble(s[4]) / 1000.0, "0.0");
                        Output("Freq.Drity:" + s[4]);
                    }
                    catch
                    {
                        Report.SetCellValue(Txsheet, row, col, "Null");
                    }
                }
                Report.SetCellProperty(Txsheet, 12, col, row, col, 9, false, 4, 1);
            }
            Report.SetCellProperty(Txsheet, 12, 2, row, col, 9, false, 4, 1);
            //Report.SetFomules(Txsheet, 12, 3, row, col,"0.0");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:HOPPing OFF; *WAI\n");

        }
        public void GetEDRrelativeTransmitPower()
        {
            int row, col;
            Excel.Worksheet Txsheet;
            List<string> listTEMP = new List<string>
            {
                listEDRtype2[listEDRtype2.Count - 1],
                listEDRtype3[listEDRtype3.Count - 1]
            };
            Output("Start EDR relative Transmit Power......");
            Txsheet = Report.AddSheet("TRM-CA-10-C");
            Report.SetCellValue(Txsheet, 1, 1, "TRM/CA/10/C EDR relative Transmit Power");
            //Report.SetCellProperty(Txsheet, 1, 1, 1, 1, 12, true, 3, -4119);
            Report.SetCellValue(Txsheet, 2, 2, "Device Name:");
            Report.SetCellValue(Txsheet, 2, 3, DeviceName);
            Report.SetCellValue(Txsheet, 3, 2, "BD Adress:");
            Report.SetCellValue(Txsheet, 3, 3, BDadress);
            Report.SetCellValue(Txsheet, 4, 2, "Burst Type:");
            Report.SetCellValue(Txsheet, 4, 3, "EDR");
            Report.SetCellValue(Txsheet, 5, 2, "Package Type:");
            Report.SetCellValue(Txsheet, 5, 3, listEDRtype2[listEDRtype2.Count - 1]);
            Report.SetCellValue(Txsheet, 5, 8, listEDRtype2[listEDRtype3.Count - 1]);
            Report.SetCellValue(Txsheet, 6, 2, "Pattern Type:");
            Report.SetCellValue(Txsheet, 6, 3, "PRBS9");
            Report.SetCellProperty(Txsheet, 3, 2, 6, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 8, 2, "Limit:");
            Report.SetCellValue(Txsheet, 8, 3, "(P_GFSK-4dB<P_DPSK<(P_GFSK+1dB)");
            Report.SetCellValue(Txsheet, 9, 2, "Limit:");
            Report.SetCellValue(Txsheet, 9, 3, "4.75<=GuardTime<=5.25");
            Report.SetCellProperty(Txsheet, 8, 2, 9, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 11, 2, "Channel");
            Report.SetCellValue(Txsheet, 11, 3, "min(P_DPSK-P_GFSK)(max power)");
            Report.SetCellValue(Txsheet, 11, 4, "max(P_DPSK-P_GFSK)(max power)");
            Report.SetCellValue(Txsheet, 11, 5, "min(P_DPSK-P_GFSK)(min power)");
            Report.SetCellValue(Txsheet, 11, 6, "max(P_DPSK-P_GFSK)(min power)");
            Report.SetCellValue(Txsheet, 11, 7, "min(GuardTime)(max power)");
            Report.SetCellValue(Txsheet, 11, 8, "max(GuardTime)(max power)");
            Report.SetCellValue(Txsheet, 11, 9, "min(GuardTime)(min power)");
            Report.SetCellValue(Txsheet, 11, 10, "max(GuardTime)(min power)");
            Report.SetCellValue(Txsheet, 11, 12, "min(P_DPSK-P_GFSK)(max power)");
            Report.SetCellValue(Txsheet, 11, 13, "max(P_DPSK-P_GFSK)(max power)");
            Report.SetCellValue(Txsheet, 11, 14, "min(P_DPSK-P_GFSK)(min power)");
            Report.SetCellValue(Txsheet, 11, 15, "max(P_DPSK-P_GFSK)(min power)");
            Report.SetCellValue(Txsheet, 11, 16, "min(GuardTime)(max power)");
            Report.SetCellValue(Txsheet, 11, 17, "max(GuardTime)(max power)");
            Report.SetCellValue(Txsheet, 11, 18, "min(GuardTime)(min power)");
            Report.SetCellValue(Txsheet, 11, 19, "max(GuardTime)(min power)");
            Report.SetCellProperty(Txsheet, 11, 2, 11, 10, 10, true, 6, -4119);
            Report.SetCellProperty(Txsheet, 11, 12, 11, 19, 10, true, 6, -4119);
            //Report.SetAutoFit(Txsheet, 11, 12, 11, 19);
            row = 11;
            col = -6;
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BTYPe EDR; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:EDRate PRBS9; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:HOPPing OFF; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:RESult:ALL OFF,OFF,OFF,OFF,OFF,OFF,OFF,ON,OFF,OFF,OFF,OFF,OFF,OFF; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:REPetition CONTinuous; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion MAX; *WAI\n");
            foreach (string varType in listTEMP)
            {
                Output("Packets type:" + varType);
                col += 9;
                row = 11;
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:EDRate " + varType + "; *WAI\n");
                foreach (int varCH in listchannel)
                {
                    Output("Channel:" + varCH.ToString());

                    row += 1;
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback " + varCH.ToString() + ",0; *WAI\n");
                    Report.SetCellValue(Txsheet, row, 2, varCH.ToString());
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                    Thread.Sleep(500);
                    Scanf(Vi, "%t", Feedback);
                    if (Feedback.ToString() != "TCON\n")
                    {
                        if (!GetVi()){return;}
                        if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                    }
                    Printf(Vi, "INITiate:BLUetooth:MEAS:MEValuation; *WAI\n");
                    Thread.Sleep(3000);
                    Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:PVTime:EDRate:MINimum?\n");
                    Failcount = 0;
                    do
                    {
                        Failcount += 1;
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                    } while (Feedback.Length == 0 && Failcount <= 20);
                    s = Feedback.ToString().Split(',');
                    if (s.Length >= 7)
                    {
                        Report.SetCellValue(Txsheet, row, col, s[5], "0.0");
                        Report.SetCellValue(Txsheet, row, col + 4, s[6], "0.0");
                    }
                    else
                    {
                        Report.SetCellValue(Txsheet, row, col, "Null");
                        Report.SetCellValue(Txsheet, row, col + 4, "Null");
                    }
                    Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:PVTime:EDRate:MAXimum?\n");
                    Failcount = 0;
                    do
                    {
                        Failcount += 1;
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                    } while (Feedback.Length == 0 && Failcount <= 20);
                    s = Feedback.ToString().Split(',');
                    if (s.Length >= 7)
                    {
                        Report.SetCellValue(Txsheet, row, col + 1, s[5], "0.0");
                        Report.SetCellValue(Txsheet, row, col + 5, s[6], "0.0");
                    }
                    else
                    {
                        Report.SetCellValue(Txsheet, row, col + 1, "Null");
                        Report.SetCellValue(Txsheet, row, col + 5, "Null");
                    }
                }
            }
            do
            {
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion DOWN; *WAI\n");
                Output("Power down...");
                Feedback.Remove(0, Feedback.Length);
                Printf(Vi, "SENSe:BLUetooth:SIGN:EUT:PCONtrol:STATe?\n");
                Thread.Sleep(500);
                Scanf(Vi, "%t", Feedback);
                s = Feedback.ToString().Split(',');
            } while (s[1] != "MIN");
            col = -6;
            foreach (string varType in listTEMP)
            {
                Output("Packets type:" + varType);
                row = 11;                
                col += 9;
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:EDRate " + varType + "; *WAI\n");
                foreach (int varCH in listchannel)
                {
                    Output("Channel:" + varCH.ToString());
                    row += 1;
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback " + varCH.ToString() + ",0; *WAI\n");
                    Report.SetCellValue(Txsheet, row, 2, varCH.ToString());
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                    Thread.Sleep(500);
                    Scanf(Vi, "%t", Feedback);
                    if (Feedback.ToString() != "TCON\n")
                    {
                        if (!GetVi()){return;}if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                    }
                    Printf(Vi, "INITiate:BLUetooth:MEAS:MEValuation; *WAI\n");
                    Thread.Sleep(3000);
                    Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:PVTime:EDRate:MINimum?\n");
                    Failcount = 0;
                    do
                    {
                        Failcount += 1;
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                    } while (Feedback.Length == 0 && Failcount <= 20);
                    s = Feedback.ToString().Split(',');
                    if (s.Length >= 7)
                    {
                        Report.SetCellValue(Txsheet, row, col + 2, s[5], "0.0");
                        Report.SetCellValue(Txsheet, row, col + 6, s[6], "0.0");
                    }
                    else
                    {
                        Report.SetCellValue(Txsheet, row, col + 2, "Null");
                        Report.SetCellValue(Txsheet, row, col + 6, "Null");
                    }
                    Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:PVTime:EDRate:MAXimum?\n");
                    Failcount = 0;
                    do
                    {
                        Failcount += 1;
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                    } while (Feedback.Length == 0 && Failcount <= 20);
                    s = Feedback.ToString().Split(',');
                    if (s.Length >= 7)
                    {
                        Report.SetCellValue(Txsheet, row, col + 3, s[5], "0.0");
                        Report.SetCellValue(Txsheet, row, col + 7, s[6], "0.0");
                    }
                    else
                    {
                        Report.SetCellValue(Txsheet, row, col + 3, "Null");
                        Report.SetCellValue(Txsheet, row, col + 7, "Null");
                    }
                }
            }

            Report.SetCellProperty(Txsheet, 12, 2, row, 10, 10, false, 4, 1);
            Report.SetCellProperty(Txsheet, 12, 12, row, 19, 10, false, 4, 1);
            //Report.SetFomules(Txsheet, 12, 3, row, 19, "0.0");

        }
        public void GetEDRcarrirFrequencyStabilityAndModulationAccurcy()
        {
            int row, col;
            Excel.Worksheet Txsheet;
            List<string> listTEMP = new List<string>
            {
                listEDRtype2[listEDRtype2.Count - 1],
                listEDRtype3[listEDRtype3.Count - 1]
            };
            Output("Start EDR carrir Frequency Stability And Modulation Accurcy...");
            Txsheet = Report.AddSheet("TRM-CA-11-C");
            Report.SetCellValue(Txsheet, 1, 1, "TRM/CA/11/C EDR Carrir Frequency Stability And Modulation Accurcy");
            //Report.SetCellProperty(Txsheet, 1, 1, 1, 1, 12, true, 3, -4119);
            Report.SetCellValue(Txsheet, 2, 2, "Device Name:");
            Report.SetCellValue(Txsheet, 2, 3, DeviceName);
            Report.SetCellValue(Txsheet, 3, 2, "BD Adress:");
            Report.SetCellValue(Txsheet, 3, 3, BDadress);
            Report.SetCellValue(Txsheet, 4, 2, "Burst Type:");
            Report.SetCellValue(Txsheet, 4, 3, "EDR");
            Report.SetCellValue(Txsheet, 5, 2, "Package Type:");
            Report.SetCellValue(Txsheet, 5, 3, listEDRtype2[listEDRtype2.Count - 1]);
            Report.SetCellValue(Txsheet, 5, 10, listEDRtype3[listEDRtype3.Count - 1]);
            Report.SetCellValue(Txsheet, 6, 2, "Pattern Type:");
            Report.SetCellValue(Txsheet, 6, 3, "PRBS9");
            Report.SetCellProperty(Txsheet, 3, 2, 6, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 8, 2, "Limit:");
            Report.SetCellValue(Txsheet, 8, 3, "|W(i)|<=75KHz");
            Report.SetCellValue(Txsheet, 8, 4, "|W(i)+W(0)|<=75KHz");
            Report.SetCellValue(Txsheet, 8, 5, "|W(0)|<=10KHz");
            Report.SetCellValue(Txsheet, 8, 6, "RMS DEVM<=0.2");
            Report.SetCellValue(Txsheet, 8, 7, "Peak DEVM<=0.35");
            Report.SetCellValue(Txsheet, 8,8, "99%DEVM<=0.3");
            Report.SetCellValue(Txsheet, 8, 10, "|W(i)|<=75KHz,,");
            Report.SetCellValue(Txsheet, 8, 11, "|W(i)+W(0)|<=75KHz");
            Report.SetCellValue(Txsheet, 8, 12, "|W(0)|<=10KHz");
            Report.SetCellValue(Txsheet, 8, 13, "RMS DEVM<=0.13");
            Report.SetCellValue(Txsheet, 8, 14, "Peak DEVM<=0.25");
            Report.SetCellValue(Txsheet, 8, 15, "99%DEVM<=0.2");

            Report.SetCellProperty(Txsheet, 8, 2, 9, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 11, 2, "Channel");
            Report.SetCellValue(Txsheet, 11, 3, "W(i) (KHz)");
            Report.SetCellValue(Txsheet, 11, 4, "[W(i)+W(0)]max (KHz)");
            Report.SetCellValue(Txsheet, 11, 5, "[W(0)]max (KHz)");
            Report.SetCellValue(Txsheet, 11, 6, "RMS DEVM");
            Report.SetCellValue(Txsheet, 11, 7, "Peak DEVM");
            Report.SetCellValue(Txsheet, 11, 8, "99% DEVM");
            Report.SetCellValue(Txsheet, 11, 10, "W(i) (KHz)");
            Report.SetCellValue(Txsheet, 11, 11, "[W(i)+W(0)]max (KHz)");
            Report.SetCellValue(Txsheet, 11, 12, "[W(0)]max (KHz)");
            Report.SetCellValue(Txsheet, 11, 13, "RMS DEVM");
            Report.SetCellValue(Txsheet, 11, 14, "Peak DEVM");
            Report.SetCellValue(Txsheet, 11, 15, "99% DEVM");

            Report.SetCellProperty(Txsheet, 11, 2, 11, 8, 10, true, 6, -4119);
            Report.SetCellProperty(Txsheet, 11, 10, 11, 15, 10, true, 6, -4119);
            //Report.SetAutoFit(Txsheet, 2, 2, 11, 15);
            row = 11;
            col = -4;
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BTYPe EDR; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:EDRate PRBS9; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:HOPPing OFF; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:RESult:ALL OFF,OFF,ON,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:REPetition CONTinuous; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion MAX; *WAI\n");
            foreach (string varType in listTEMP)
            {
                Output("Packets type:"+varType);
                col += 7;
                row = 11;
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:EDRate " + varType + "; *WAI\n");
                foreach (int varCH in listchannel)
                {
                    Output("Channel:" + varCH.ToString());
                    row += 1;
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback " + varCH.ToString() + ",0; *WAI\n");
                    Report.SetCellValue(Txsheet, row, 2, varCH.ToString());
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                    Thread.Sleep(500);
                    Scanf(Vi, "%t", Feedback);
                    if (Feedback.ToString() != "TCON\n")
                    {
                        if (!GetVi()){return;}
                        if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                    }
                    Printf(Vi, "INITiate:BLUetooth:MEAS:MEValuation; *WAI\n");
                    Thread.Sleep(5000);
                    Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:MODulation:EDRate:AVERage?\n");
                    Failcount = 0;
                    do
                    {
                        Failcount += 1;
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                    } while (Feedback.Length == 0 && Failcount <= 20);
                    s = Feedback.ToString().Split(',');
                    if (s.Length >= 8)
                    {
                        Report.SetCellValue(Txsheet, row, col, s[2],"0.0");
                        Report.SetCellValue(Txsheet, row, col + 1, s[3], "0.0");
                        Report.SetCellValue(Txsheet, row, col + 2, s[4], "0.0");
                        Report.SetCellValue(Txsheet, row, col + 3, s[5], "0.000");
                        Report.SetCellValue(Txsheet, row, col + 4, s[6], "0.000");
                        Report.SetCellValue(Txsheet, row, col + 5, s[7],"0.000");
                        Output("success");
                    }
                    else
                    {
                        Report.SetCellValue(Txsheet, row, col, "Null");
                        Report.SetCellValue(Txsheet, row, col + 1, "Null");
                        Report.SetCellValue(Txsheet, row, col + 2, "Null");
                        Report.SetCellValue(Txsheet, row, col + 3, "Null");
                        Report.SetCellValue(Txsheet, row, col + 4, "Null");
                        Report.SetCellValue(Txsheet, row, col + 5, "Null");
                        Output("fail");
                    }
                }
            }
            Report.SetCellProperty(Txsheet, 12, 2, row, 8, 10, false, 4, 1);
            Report.SetCellProperty(Txsheet, 12, 10, row, 15, 10, false, 4, 1);
           // Report.SetFomules(Txsheet, 13, 3, row, 15,"0.00");
        }
        public void GetEDRdifferentialPhaseEncoding()
        {
            int row, col;
            Excel.Worksheet Txsheet;
            List<string> listTEMP = new List<string>
            {
                "E21P",
                "E31P"
            };
            Output("Start EDR Differential Phase Encoding......");
            Txsheet = Report.AddSheet("TRM-CA-12-C");
            Report.SetCellValue(Txsheet, 1, 1, "TRM/CA/12/C EDR Differential Phase Encoding");
            //Report.SetCellProperty(Txsheet, 1, 1, 1, 1, 12, true, 3, -4119);
            Report.SetCellValue(Txsheet, 2, 2, "Device Name:");
            Report.SetCellValue(Txsheet, 2, 3, DeviceName);
            Report.SetCellValue(Txsheet, 3, 2, "BD Adress:");
            Report.SetCellValue(Txsheet, 3, 3, BDadress);
            Report.SetCellValue(Txsheet, 4, 2, "Burst Type:");
            Report.SetCellValue(Txsheet, 4, 3, "EDR");
            Report.SetCellValue(Txsheet, 5, 2, "Package Type:");
            Report.SetCellValue(Txsheet, 5, 3, "2DH1");
            Report.SetCellValue(Txsheet, 5, 5, "3DH1");
            Report.SetCellValue(Txsheet, 6, 2, "Pattern Type:");
            Report.SetCellValue(Txsheet, 6, 3, "PRBS9");
            Report.SetCellProperty(Txsheet, 2, 2, 6, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 8, 2, "Limit:");
            Report.SetCellValue(Txsheet, 8, 3, "zero errors  in 99% of the packets");
            Report.SetCellProperty(Txsheet, 8, 2, 8, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 11, 2, "Channel");
            Report.SetCellValue(Txsheet, 11, 3, "Packets with 0 bit error (%)");
            Report.SetCellValue(Txsheet, 11, 5, "Packets with 0 bit error (%)");


            Report.SetCellProperty(Txsheet, 11, 2, 11, 3, 10, true, 6, -4119);
            Report.SetCellProperty(Txsheet, 11, 5, 11, 5, 10, true, 6, -4119);
            //Report.SetAutoFit(Txsheet, 2, 2, 11, 5);
            row = 11;
            col = 1;
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BTYPe EDR; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:EDRate PRBS9; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:HOPPing OFF; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe TXTest; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:RESult:ALL OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,ON; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:REPetition CONTinuous; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion MAX; *WAI\n");
            foreach (string varType in listTEMP)
            {
                Output("Packets type:" + varType);
                col += 2;
                row = 11;
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:EDRate " + varType + "; *WAI\n");
                foreach (int varCH in listchannel)
                {
                    Output("Channel:" + varCH.ToString());
                    row += 1;
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:TXTest " + varCH.ToString() + "; *WAI\n");
                    Report.SetCellValue(Txsheet, row, 2, varCH.ToString());
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                    Thread.Sleep(500);
                    Scanf(Vi, "%t", Feedback);
                    if (Feedback.ToString() != "TCON\n")
                    {
                        if (!GetVi()){return;}
                        if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                    }
                    Printf(Vi, "INITiate:BLUetooth:MEAS:MEValuation; *WAI\n");
                    Thread.Sleep(3000);
                    Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:PENCoding:EDRate:CURRent:C?\n");
                    Failcount = 0;
                    do
                    {
                        Failcount += 1;
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                    } while (Feedback.Length == 0 && Failcount <= 20);
                    s = Feedback.ToString().Split(',');
                    if (s.Length >= 4)
                    {
                        Report.SetCellValue(Txsheet, row, col, s[3],"0.000");
                    }
                    else
                    {
                        Report.SetCellValue(Txsheet, row, col, "Null");

                    }
                }
            }
            Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
            Report.SetCellProperty(Txsheet, 12, 2, row, 3, 10, false, 4, 1);
            Report.SetCellProperty(Txsheet, 12, 5, row, 5, 10, false, 4, 1);
            //Report.SetFomules(Txsheet, 12, 3, row, 5, "0.000");
        }
        /*
        public void GetEDRin_BandSpuriousEmission()
        {
            int row, col,TempI;
            Excel.Worksheet Txsheet;
            List<string> listTEMP = new List<string>();
            listTEMP.Add(listEDRtype2[listEDRtype2.Count - 1]);
            listTEMP.Add(listEDRtype3[listEDRtype3.Count - 1]);
            Txsheet = Report.AddSheet("TRM-CA-13-C");
            Report.SetCellValue(Txsheet, 1, 1, "TRM/CA/13/C EDR In-Band Spurious Emission");
            //Report.SetCellProperty(Txsheet, 1, 1, 1, 1, 12, true, 3, -4119);
            Report.SetCellValue(Txsheet, 3, 2, "Device Name:");
            Report.SetCellValue(Txsheet, 3, 3, DeviceName);
            Report.SetCellValue(Txsheet, 4, 2, "Burst Type:");
            Report.SetCellValue(Txsheet, 4, 3, "EDR");
            Report.SetCellValue(Txsheet, 5, 2, "Package Type:");
            Report.SetCellValue(Txsheet, 5, 3, listEDRtype2[listEDRtype2.Count - 1]);
            Report.SetCellValue(Txsheet, 5, 5, listEDRtype3[listEDRtype3.Count - 1]);
            Report.SetCellValue(Txsheet, 6, 2, "Pattern Type:");
            Report.SetCellValue(Txsheet, 6, 3, "PRBS9");
            Report.SetCellProperty(Txsheet, 3, 2, 6, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 8, 2, "Limit:");
            Report.SetCellValue(Txsheet, 8, 3, "PTX-26dB(f)<PTXref-26dB for |M-N|=1");
            Report.SetCellValue(Txsheet, 9, 3, "PTX(f)<–20dBm for |M-N|=2");
            Report.SetCellValue(Txsheet, 10, 3, "PTX(f)<–40dBm for |M-N|>=3");
            Report.SetCellProperty(Txsheet, 8, 2, 8, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 11, 2, "Channel");
            Report.SetCellValue(Txsheet, 11, 3, "Packets with 0 bit error (%)");
            Report.SetCellValue(Txsheet, 11, 5, "Packets with 0 bit error (%)");


            Report.SetCellProperty(Txsheet, 11, 2, 11, 3, 10, true, 6, -4119);
            Report.SetCellProperty(Txsheet, 11, 5, 11, 5, 10, true, 6, -4119);
            row = 11;
            col = 1;
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BTYPe EDR; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:EDRate PRBS9; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:HOPPing OFF; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:RESult:ALL OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,OFF,ON,OFF,OFF,OFF; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:REPetition CONTinuous; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:SGACp:EDRate:MEASurement:MODE CH21; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion MAX; *WAI\n");
            foreach (string varType in listTEMP)
            {
                col += 3;
                row = 11;
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:EDRate " + varType + "; *WAI\n");
                foreach (int varCH in listACP)
                {
                    col += 1;
                    row = 12;
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                    Scanf(Vi, "%t", Feedback);
                    if (Feedback.ToString() != "TCON\n")
                    {
                        if (!GetVi()){return;}if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                    }
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback " + varCH.ToString() + ",0; *WAI\n");
                    //Report.SetCellValue(Txsheet, row, 2, varCH.ToString());
                    Printf(Vi, "INITiate:BLUetooth:MEAS:MEValuation; *WAI\n");
                    Thread.Sleep(3000);

                    Printf(Vi, "FETCh:BLUetooth:MEAS:MEValuation:SGACp:EDRate?\n");
                    Failcount = 0;
                    do
                    {
                        Failcount += 1;
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                    } while (Feedback.Length == 0 && Failcount <= 20);
                    s = Feedback.ToString().Split(',');
                    Report.SetCellValue(Txsheet, row, 2, "Number of exceptions:");
                    if (s.Length >= 82)
                    {
                        Report.SetCellValue(Txsheet, row, col, s[3]);
                    }
                    else
                    {
                        Report.SetCellValue(Txsheet, row, col, "Null");
                        break;
                    }

                    row = 13;
                    for (TempI = 0; TempI < 79; TempI++, row++)
                    {
                        Report.SetCellValue(Txsheet, row, 2, TempI.ToString());
                        Report.SetCellValue(Txsheet, row, col, s[TempI + 4]);
                    }
                }
                Report.SetCellProperty(Txsheet, 12, 2, row - 1, 5, 9, false, 4, 1);
            }

        }

        */
        public void GetSensitivityMultiSlot()
        {
            int row, col,Packets,TempI,FailC;
            bool IsEnd;
            double TxPower, Gep;
            Excel.Worksheet Txsheet;
            Output("Start Sensitivity Multi-Slot......");
            List<int> PacketsNUM = new List<int>
            {
                7200,
                1100,
                600
            };
            List<string> listTEMP = new List<string>
            {
                "DH1",
                "DH3",
                "DH5"
            };
            Txsheet = Report.AddSheet("RCV-CA-01-02-C");
            Report.SetCellValue(Txsheet, 1, 1, "RCV/CA/01-02/C Sensitivity slot packets");
            //Report.SetCellProperty(Txsheet, 1, 1, 1, 1, 12, true, 3, -4119);
            Report.SetCellValue(Txsheet, 2, 2, "Device Name:");
            Report.SetCellValue(Txsheet, 2, 3, DeviceName);
            Report.SetCellValue(Txsheet, 3, 2, "BD Adress:");
            Report.SetCellValue(Txsheet, 3, 3, BDadress);
            Report.SetCellValue(Txsheet, 4, 2, "Burst Type:");
            Report.SetCellValue(Txsheet, 4, 3, "BR");
            Report.SetCellValue(Txsheet, 5, 2, "Sensitivity Type:");
            Report.SetCellValue(Txsheet, 5, 3, "Single slot");
            Report.SetCellValue(Txsheet, 5, 6, "Multi slot");
            Report.SetCellValue(Txsheet, 6, 2, "Package Type:");
            Report.SetCellValue(Txsheet, 6, 3, "DH1");
            Report.SetCellValue(Txsheet, 6, 8, "DH3");
            Report.SetCellValue(Txsheet, 6, 13, "DH5");
            Report.SetCellValue(Txsheet, 7, 2, "Pattern Type:");
            Report.SetCellValue(Txsheet, 7, 3, "PRBS9");
            Report.SetCellValue(Txsheet, 8, 2, "Dirty TX:");
            Report.SetCellValue(Txsheet, 8, 3, "ON");
            Report.SetCellValue(Txsheet, 9, 2, "Hopping:");
            Report.SetCellValue(Txsheet, 9, 3, "ON");
            Report.SetCellProperty(Txsheet, 2, 2, 9, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 11, 2, "Limit:");
            Report.SetCellValue(Txsheet, 11, 3, "BER ≤ 0.1%(minimum number of samples, 1 600 000 returned payload bits.)");
            Report.SetCellProperty(Txsheet, 11, 2, 11, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 13, 2, "Channel");
            Report.SetCellValue(Txsheet, 13, 3, "BER(%)");
            Report.SetCellValue(Txsheet, 13, 4, "RX level(dBm）");
            Report.SetCellValue(Txsheet, 13, 5, "BER(%)");
            Report.SetCellValue(Txsheet, 13, 6, "RX level(dBm)");
            Report.SetCellValue(Txsheet, 13, 8, "BER(%)");
            Report.SetCellValue(Txsheet, 13, 9, "RX level(dBm）");
            Report.SetCellValue(Txsheet, 13, 10, "BER(%)");
            Report.SetCellValue(Txsheet, 13, 11, "RX level(dBm)");
            Report.SetCellValue(Txsheet, 13, 13, "BER(%)");
            Report.SetCellValue(Txsheet, 13, 14, "RX level(dBm）");
            Report.SetCellValue(Txsheet, 13, 15, "BER(%)");
            Report.SetCellValue(Txsheet, 13, 16, "RX level(dBm)");
            Report.SetCellProperty(Txsheet, 13, 2, 13, 6, 10, true, 6, -4119);
            Report.SetCellProperty(Txsheet, 13, 8, 13, 11, 10, true, 6, -4119);
            Report.SetCellProperty(Txsheet, 13, 13, 13, 16, 10, true, 6, -4119);
            //Report.SetAutoFit(Txsheet, 2, 2, 13, 6);

            TxPower = -90.0;
            TempI = 0;
            IsEnd = false;
            Packets = 14200;
            row = 13;
            col = -2;
            Gep = -0.5;
            Printf(Vi, "CONFigure:BLUetooth:SIGN:OPMode RFT;; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BTYPe BR; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:HWINterface1 NONE; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:BRATe DH1; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion MAX; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:BRATe PRBS9; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:HOPPing ON; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:DTX ON; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:DTX:MODE:BRATe SPEC; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RXQuality:REPetition SINGleshot; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RXQuality:PACKets 2000; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:LEVel -70; *WAI\n");
            Printf(Vi, "; *WAI\n");
            Printf(Vi, "; *WAI\n");
            foreach(string varType in listBRtype)
            {
                Output("Packets type:" + varType);
                col += 5;
                row = 13;
                Printf(Vi, "CONFigure:BLUetooth:SIGN:RXQuality:PACKets "+PacketsNUM[TempI].ToString()+"; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:BRATe "+varType+"; *WAI\n");
                
                foreach (int varCH in listchannel)
                {
                    Output("Channel:" + varCH.ToString());
                    row += 1;
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                    Scanf(Vi, "%t", Feedback);
                    if (Feedback.ToString() != "TCON\n")
                    {
                        if (!GetVi()){return;}
                        if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                    }
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback 0," + varCH.ToString() + "; *WAI\n");
                    Report.SetCellValue(Txsheet, row, 2, varCH.ToString());
                    Printf(Vi, "READ:BLUetooth:SIGN:RXQuality:BER?\n");
                    Failcount = 0;
                    do
                    {
                        Failcount += 1;
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                        s = Feedback.ToString().Split(',');
                    } while (Feedback.Length == 0 && Failcount <= 80);
                    Report.SetCellValue(Txsheet, row, col, s[1], "0.000");
                    Report.SetCellValue(Txsheet, row, col + 1, "-70");
                }
                row = 13;
                Packets = PacketsNUM[TempI]/4;
                Printf(Vi, "CONFigure:BLUetooth:SIGN:RXQuality:PACKets " + Packets.ToString() + "; *WAI\n");
                foreach (int varCH in listchannel)
                {
                    Output("Channel:" + varCH.ToString());
                    row += 1;
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                    Scanf(Vi, "%t", Feedback);
                    if (Feedback.ToString() != "TCON\n")
                    {
                        if (!GetVi()){return;}if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                    }
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback 0," + varCH.ToString() + "; *WAI\n");
                    Failcount = 0;
                    try
                    {
                        IsEnd = false;
                        do
                        {                            
                            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:LEVel " + TxPower.ToString() + "; *WAI\n");
                            Feedback.Remove(0, Feedback.Length);
                            Printf(Vi, "*OPC?\n");
                            Thread.Sleep(500);
                            Scanf(Vi, "%t", Feedback);
                            Feedback.Remove(0, Feedback.Length);
                            Printf(Vi, "READ:BLUetooth:SIGN:RXQuality:BER?\n");
                            FailC = 0;
                            do
                            {
                                FailC += 1;
                                Feedback.Remove(0, Feedback.Length);
                                Thread.Sleep(500);
                                Scanf(Vi, "%t", Feedback);
                                s = Feedback.ToString().Split(',');
                            } while (Feedback.Length == 0 && FailC <= 80);
                            s = Feedback.ToString().Split(',');
                            if (Convert.ToDouble(s[1]) >= 0.1) { Gep = 0.1; IsEnd = true; }
                            Failcount += 1;
                            TxPower += Gep;
                            if (Failcount >= 50||TxPower>-70) { break; }
                        } while (!(Convert.ToDouble(s[1]) <= 0.1 && IsEnd));
                        IsEnd = false;
                        Gep = -0.5;
                        Report.SetCellValue(Txsheet, row, col + 3, TxPower.ToString());
                        Report.SetCellValue(Txsheet, row, col + 2, s[1], "0.000");
                    }
                    catch
                    {
                        Report.SetCellValue(Txsheet, row, col+3, "Null");
                        Report.SetCellValue(Txsheet, row, col + 2, "Null");
                        IsEnd = false;
                        Gep = -0.5;
                        continue;
                    }

                }
                TempI++;
            }       
            Report.SetCellProperty(Txsheet, 14, 2, row, 6, 9, false, 4, 1);
            Report.SetCellProperty(Txsheet, 14, 8, row, 11, 9, false, 4, 1);
            Report.SetCellProperty(Txsheet, 14, 13, row, 16, 9, false, 4, 1);
          //Report.SetFomules(Txsheet, 14, 3, row, 16,"0.000");
        }
        public void GetMaximumInputLevel()
        {
            int row, col;
            Excel.Worksheet Txsheet;
            Output("Start Maximum Input Level......");
            Txsheet = Report.AddSheet("RCV-CA-06-C");
            Report.SetCellValue(Txsheet, 1, 1, "RCV/CA/06/C MaximumInputLevel");
            //Report.SetCellProperty(Txsheet, 1, 1, 1, 1, 12, true, 3, -4119);
            Report.SetCellValue(Txsheet, 2, 2, "Device Name:");
            Report.SetCellValue(Txsheet, 2, 3, DeviceName);
            Report.SetCellValue(Txsheet, 3, 2, "BD Adress:");
            Report.SetCellValue(Txsheet, 3, 3, BDadress);
            Report.SetCellValue(Txsheet, 4, 2, "Burst Type:");
            Report.SetCellValue(Txsheet, 4, 3, "BR");
            Report.SetCellValue(Txsheet, 5, 2, "Package Type:");
            Report.SetCellValue(Txsheet, 5, 3, "DH1");
            Report.SetCellValue(Txsheet, 6, 2, "Pattern Type:");
            Report.SetCellValue(Txsheet, 6, 3, "PRBS9");
            Report.SetCellValue(Txsheet, 7, 2, "RX Level:");
            Report.SetCellValue(Txsheet, 7, 3, "-20dBm");
            Report.SetCellValue(Txsheet, 8, 2, "Dirty TX:");
            Report.SetCellValue(Txsheet, 8, 3, "OFF");
            Report.SetCellValue(Txsheet, 9, 2, "Hopping:");
            Report.SetCellValue(Txsheet, 9, 3, "OFF");
            Report.SetCellProperty(Txsheet, 2, 2, 9, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 11, 2, "Limit:");
            Report.SetCellValue(Txsheet, 11, 3, "BER ≤ 0.1%(minimum number of samples, 1 600 000 returned payload bits.)");
            Report.SetCellProperty(Txsheet, 11, 2, 11, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 13, 2, "Channel");
            Report.SetCellValue(Txsheet, 13, 3, "BER");
            Report.SetCellValue(Txsheet, 13, 4, "RX level");
            Report.SetCellProperty(Txsheet, 13, 2, 13, 4, 10, true, 6, -4119);
            row = 13;
            col = 3;
            Printf(Vi, "CONFigure:BLUetooth:SIGN:OPMode RFT;; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BTYPe BR; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:HWINterface1 NONE; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:BRATe DH1; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion MAX; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:BRATe PRBS9; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:HOPPing OFF; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:DTX OFF; *WAI\n");
            //Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:DTX:MODE:BRATe SPEC; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RXQuality:REPetition SINGleshot; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RXQuality:RXQuality:PACKets 7400; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:LEVel -20; *WAI\n");
            Printf(Vi, "; *WAI\n");
            Printf(Vi, "; *WAI\n");
            foreach (int varCH in listchannel)
            {
                Output("Channel:" + varCH.ToString());
                row += 1;
                Feedback.Remove(0, Feedback.Length);
                Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                Scanf(Vi, "%t", Feedback);
                if (Feedback.ToString() != "TCON\n")
                {
                    if (!GetVi()){return;}
                    if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                }
                Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback 0," + varCH.ToString() + "; *WAI\n");
                Report.SetCellValue(Txsheet, row, 2, varCH.ToString());
                //Printf(Vi, "READ:BLUetooth:SIGN:RXQuality:BER?\n");
                Feedback.Remove(0, Feedback.Length);
                Printf(Vi, "READ:BLUetooth:SIGN:RXQuality:BER?\n");
                Thread.Sleep(500);
                Scanf(Vi, "%t", Feedback);
                s = Feedback.ToString().Split(',');
                Report.SetCellValue(Txsheet, row, col + 1, "-20");
                Report.SetCellValue(Txsheet, row, col, s[1], "0.00000");
            }
            Report.SetCellProperty(Txsheet, 14, 2, row, 4, 9, false, 4, 1);
        }
        public void GetEDRSensitivity()
        {
            int row, col,Packets,TempI,FailC;
            Excel.Worksheet Txsheet;
            List<int> PacketsNUM = new List<int>
            {
                3000,
                2000
            };
            double TxPower,Gep;
            //string ExpPower;
            bool IsEnd;
            List<string> listTEMP = new List<string>
            {
                listEDRtype2[listEDRtype2.Count - 1],
                listEDRtype3[listEDRtype3.Count - 1]
            };
            Output("Start EDR Sensitivity......");
            Txsheet = Report.AddSheet("RCV-CA-07-C");
            Report.SetCellValue(Txsheet, 1, 1, "RCV/CA/07/C EDR Sensitivity");
            //Report.SetCellProperty(Txsheet, 1, 1, 1, 1, 12, true, 3, -4119);
            Report.SetCellValue(Txsheet, 2, 2, "Device Name:");
            Report.SetCellValue(Txsheet, 2, 3, DeviceName);
            Report.SetCellValue(Txsheet, 3, 2, "BD Adress:");
            Report.SetCellValue(Txsheet, 3, 3, BDadress);
            Report.SetCellValue(Txsheet, 4, 2, "Burst Type:");
            Report.SetCellValue(Txsheet, 4, 3, "EDR");
            Report.SetCellValue(Txsheet, 5, 2, "Package Type:");
            Report.SetCellValue(Txsheet, 5, 3, listEDRtype2[listEDRtype2.Count - 1]);
            Report.SetCellValue(Txsheet, 5, 8, listEDRtype3[listEDRtype3.Count - 1]);
            Report.SetCellValue(Txsheet, 6, 2, "Pattern Type:");
            Report.SetCellValue(Txsheet, 6, 3, "PRBS9");
            Report.SetCellValue(Txsheet, 7, 2, "RX Level:");
            Report.SetCellValue(Txsheet, 7, 3, "-70dBm");
            Report.SetCellValue(Txsheet, 8, 2, "Dirty TX:");
            Report.SetCellValue(Txsheet, 8, 3, "ON");
            Report.SetCellValue(Txsheet, 9, 2, "Whitening:");
            Report.SetCellValue(Txsheet, 9, 3, "ON");
            Report.SetCellProperty(Txsheet, 2, 2, 9, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 11, 2, "Limit:");
            Report.SetCellValue(Txsheet, 11, 3, "BER < 7•10E-5 after 1 600 000 bits or BER < 10E-4 after 16 000 000 bits");
            Report.SetCellProperty(Txsheet, 11, 2, 11, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 13, 2, "Channel");
            Report.SetCellValue(Txsheet, 13, 3, "BER 16 000 000");
            Report.SetCellValue(Txsheet, 13, 4, "RX level");
            Report.SetCellValue(Txsheet, 13, 5, "BER");
            Report.SetCellValue(Txsheet, 13, 6, "RX level");
            Report.SetCellValue(Txsheet, 13, 8, "BER 16 000 000");
            Report.SetCellValue(Txsheet, 13, 9, "RX level");
            Report.SetCellValue(Txsheet, 13, 10, "BER");
            Report.SetCellValue(Txsheet, 13, 11, "RX level");
            Report.SetCellProperty(Txsheet, 13, 2, 13, 6, 10, true, 6, -4119);
            Report.SetCellProperty(Txsheet, 13, 8, 13, 11, 10, true, 6, -4119);
            //Report.SetAutoFit(Txsheet, 2, 2, 13, 7);
            Packets = 10000;
            
            IsEnd = false;
            TempI = 0;
            Gep = -0.5;
            row = 13;
            col = -2;
            Printf(Vi, "CONFigure:BLUetooth:SIGN:OPMode RFT;; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BTYPe EDR; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:HWINterface1 NONE; *WAI\n");
            //Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:BRATe DH1; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion MAX; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:EDRate PRBS9; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:HOPPing OFF; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:WHITening ON; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:DTX ON; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:DTX:MODE:EDRate SPEC; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RXQuality:REPetition SINGleshot; *WAI\n");
            //Printf(Vi, "CONFigure:BLUetooth:SIGN:RXQuality:PACKets 7400; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:LEVel -70; *WAI\n");
            foreach(string varType in listTEMP)
            {
                Output("Channel:" + varType);
                TxPower = -85.0;
                row = 13;
                col += 5;
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:EDRate "+varType+"; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:RXQuality:PACKets " + PacketsNUM[TempI].ToString() + "; *WAI\n");
                
                foreach (int varCH in listchannel)
                {
                    Output("Channel:" + varCH.ToString());
                    row += 1;
                    
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                    Scanf(Vi, "%t", Feedback);
                    if (Feedback.ToString() != "TCON\n")
                    {
                        if (!GetVi()){return;}
                        if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                    }
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback 0," + varCH.ToString() + "; *WAI\n");
                    Report.SetCellValue(Txsheet, row, 2, varCH.ToString());
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "READ:BLUetooth:SIGN:RXQuality:BER?\n");
                    Failcount = 0;
                    do
                    {
                        Failcount += 1;
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                    } while (Feedback.Length == 0 && Failcount <= 80);
                    s = Feedback.ToString().Split(',');
                    Report.SetCellValue(Txsheet, row, col + 1, "-70");
                    Report.SetCellValue(Txsheet, row, col, s[1], "0.00000");
                }
                row = 13;
                Packets = PacketsNUM[TempI]/40;
                Printf(Vi, "CONFigure:BLUetooth:SIGN:RXQuality:PACKets " + Packets.ToString() + "; *WAI\n");
                foreach (int varCH in listchannel)
                {
                    row += 1;
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                    Scanf(Vi, "%t", Feedback);
                    if (Feedback.ToString() != "TCON\n")
                    {
                        if (!GetVi()){return;}
                        if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                    }
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback 0," + varCH.ToString() + "; *WAI\n");
                    
                    Failcount = 0;
                    try
                    {
                        do
                        {
                            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:LEVel " + TxPower.ToString() + "; *WAI\n");
                            Printf(Vi, "*OPC?\n");
                            Feedback.Remove(0, Feedback.Length);
                            Thread.Sleep(500);
                            Scanf(Vi, "%t", Feedback);
                            Feedback.Remove(0, Feedback.Length);
                            Printf(Vi, "READ:BLUetooth:SIGN:RXQuality:BER?\n");
                            FailC = 0;
                            do
                            {
                                FailC += 1;
                                Feedback.Remove(0, Feedback.Length);
                                Thread.Sleep(500);
                                Scanf(Vi, "%t", Feedback);
                                s = Feedback.ToString().Split(',');
                            } while (Feedback.Length == 0 && FailC <= 40);
                            s = Feedback.ToString().Split(',');
                            if (Convert.ToDouble(s[1]) >= 0.0001) { Gep = 0.1; IsEnd = true; }
                            Failcount += 1;
                            TxPower += Gep;
                            
                            if (Failcount >= 50 || TxPower > -70) { break; }
                        } while (!(Convert.ToDouble(s[1]) <= 0.0001 && IsEnd));
                        IsEnd = false;
                        Gep = -0.5;
                        Report.SetCellValue(Txsheet, row, col + 3, TxPower.ToString());
                        Report.SetCellValue(Txsheet, row, col + 2, s[1], "0.00000");
                        //Report.SetFomules(Txsheet,row, col + 2, row, col + 2, "0.00000");
                    }
                    catch
                    {
                        Report.SetCellValue(Txsheet, row, col + 3, "Null");
                        Report.SetCellValue(Txsheet, row, col + 2, "Null");
                        IsEnd = false;
                        Gep = -0.5;
                        continue;
                    }
                    
                }
                TempI++;
            }
            
            //Printf(Vi, "; *WAI\n");
            
            Report.SetCellProperty(Txsheet, 14, 2, row, 6, 9, false, 4, 1);
            Report.SetCellProperty(Txsheet, 14, 8, row, 11, 9, false, 4, 1);
            //Report.SetFomules(Txsheet, 14, 3, row, 3, "0.00000");

        }
        public void GetEDRBERFloorPerformance()
        {
            int row, col,TempI;
            Excel.Worksheet Txsheet;
            List<string> PacketsNUM = new List<string>
            {
                "1500",
                "1000"
            };
            List<string> listTEMP = new List<string>
            {
                listEDRtype2[listEDRtype2.Count - 1],
                listEDRtype3[listEDRtype3.Count - 1]
            };
            Output("Start EDR BER Floor Performance......");
            Txsheet = Report.AddSheet("RCV-CA-08-C");
            Report.SetCellValue(Txsheet, 1, 1, "RCV/CA/08/C EDR BER Floor Performance");
            //Report.SetCellProperty(Txsheet, 1, 1, 1, 1, 12, true, 3, -4119);
            Report.SetCellValue(Txsheet, 2, 2, "Device Name:");
            Report.SetCellValue(Txsheet, 2, 3, DeviceName);
            Report.SetCellValue(Txsheet, 3, 2, "BD Adress:");
            Report.SetCellValue(Txsheet, 3, 3, BDadress);
            Report.SetCellValue(Txsheet, 4, 2, "Burst Type:");
            Report.SetCellValue(Txsheet, 4, 3, "EDR");
            Report.SetCellValue(Txsheet, 5, 2, "Package Type:");
            Report.SetCellValue(Txsheet, 5, 3, listEDRtype2[listEDRtype2.Count - 1]);
            Report.SetCellValue(Txsheet, 5, 6, listEDRtype3[listEDRtype3.Count - 1]);
            Report.SetCellValue(Txsheet, 6, 2, "Pattern Type:");
            Report.SetCellValue(Txsheet, 6, 3, "PRBS9");
            Report.SetCellValue(Txsheet, 7, 2, "RX Level:");
            Report.SetCellValue(Txsheet, 7, 3, "-60dBm");
            Report.SetCellValue(Txsheet, 8, 2, "Dirty TX:");
            Report.SetCellValue(Txsheet, 8, 3, "OFF");
            Report.SetCellValue(Txsheet, 9, 2, "Whitening:");
            Report.SetCellValue(Txsheet, 9, 3, "ON");
            Report.SetCellProperty(Txsheet, 2, 2, 9, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 11, 2, "Limit:");
            Report.SetCellValue(Txsheet, 11, 3, "BER < 7•10-6 after 8 000 000 bits or BER < 10-5 after 160 000 000 bits");
            Report.SetCellProperty(Txsheet, 11, 2, 11, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 13, 2, "Channel");
            Report.SetCellValue(Txsheet, 13, 3, "BER 8 000 000");
            Report.SetCellValue(Txsheet, 13, 4, "RX level(dBm)");
            Report.SetCellValue(Txsheet, 13, 6, "BER 8 000 000");
            Report.SetCellValue(Txsheet, 13, 7, "RX level(dBm)");
            Report.SetCellProperty(Txsheet, 13, 2, 13, 4, 10, true, 6, -4119);
            Report.SetCellProperty(Txsheet, 13, 6, 13, 7, 10, true, 6, -4119);
            //Report.SetAutoFit(Txsheet, 2, 2, 13, 2);
            //Report.SetAutoFit(Txsheet, 13, 3, 13, 7);
            row = 13;
            col = 0;
            TempI = 0;
            Printf(Vi, "CONFigure:BLUetooth:SIGN:OPMode RFT;; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BTYPe EDR; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:HWINterface1 NONE; *WAI\n");
            //Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:BRATe DH1; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion MAX; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:EDRate PRBS9; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:HOPPing OFF; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:WHITening ON; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:DTX OFF; *WAI\n");
            //Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:DTX:MODE:EDRate SPEC; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RXQuality:REPetition SINGleshot; *WAI\n");
            
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:LEVel -60; *WAI\n");
            foreach (string varType in listTEMP)
            {
                Output("Packets type:" + varType);
                col += 3;
                row = 13;
                Printf(Vi, "CONFigure:BLUetooth:SIGN:RXQuality:PACKets " + PacketsNUM[TempI].ToString() + "; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:EDRate " + varType + "; *WAI\n");
                TempI++;
                foreach (int varCH in listchannel)
                {
                    Output("Channel:" + varCH.ToString());
                    row += 1;
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                    Scanf(Vi, "%t", Feedback);
                    if (Feedback.ToString() != "TCON\n")
                    {
                        if (!GetVi()){return;}
                        if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                    }
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback 0," + varCH.ToString() + "; *WAI\n");
                    Report.SetCellValue(Txsheet, row, 2, varCH.ToString());
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "READ:BLUetooth:SIGN:RXQuality:BER?\n");
                    Failcount = 0;
                    do
                    {
                        Failcount += 1;
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                    } while (Feedback.Length == 0 && Failcount <= 80);
                    s = Feedback.ToString().Split(',');
                    Report.SetCellValue(Txsheet, row, col + 1, "-60");
                    Report.SetCellValue(Txsheet, row, col, s[1], "0.0000000");
                }
            }

            //Printf(Vi, "; *WAI\n");

            Report.SetCellProperty(Txsheet, 14, 2, row, 4, 9, false, 4, 1);
            Report.SetCellProperty(Txsheet, 14, 6, row, 7, 9, false, 4, 1);


        }
        public void GetEDRMaximumInputLevel()
        {
            int row, col, TempI;
            Excel.Worksheet Txsheet;
            List<string> PacketsNUM = new List<string>
            {
                "1500",
                "1000"
            };
            List<string> listTEMP = new List<string>
            {
                listEDRtype2[listEDRtype2.Count - 1],
                listEDRtype3[listEDRtype3.Count - 1]
            };
            Output("Start EDR Maximum Input Level......");
            Txsheet = Report.AddSheet("RCV-CA-10-C");
            Report.SetCellValue(Txsheet, 1, 1, "RCV/CA/10/C EDR BER Floor Performance");
            //Report.SetCellProperty(Txsheet, 1, 1, 1, 1, 12, true, 3, -4119);
            Report.SetCellValue(Txsheet, 2, 2, "Device Name:");
            Report.SetCellValue(Txsheet, 2, 3, DeviceName);
            Report.SetCellValue(Txsheet, 3, 2, "BD Adress:");
            Report.SetCellValue(Txsheet, 3, 3, BDadress);
            Report.SetCellValue(Txsheet, 4, 2, "Burst Type:");
            Report.SetCellValue(Txsheet, 4, 3, "EDR");
            Report.SetCellValue(Txsheet, 5, 2, "Package Type:");
            Report.SetCellValue(Txsheet, 5, 3, listEDRtype2[listEDRtype2.Count - 1]);
            Report.SetCellValue(Txsheet, 5, 6, listEDRtype3[listEDRtype3.Count - 1]);
            Report.SetCellValue(Txsheet, 6, 2, "Pattern Type:");
            Report.SetCellValue(Txsheet, 6, 3, "PRBS9");
            Report.SetCellValue(Txsheet, 7, 2, "RX Level:");
            Report.SetCellValue(Txsheet, 7, 3, "-20dBm");
            Report.SetCellValue(Txsheet, 8, 2, "Dirty TX:");
            Report.SetCellValue(Txsheet, 8, 3, "OFF");
            Report.SetCellValue(Txsheet, 9, 2, "Whitening:");
            Report.SetCellValue(Txsheet, 9, 3, "ON");
            Report.SetCellProperty(Txsheet, 2, 2, 9, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 11, 2, "Limit:");
            Report.SetCellValue(Txsheet, 11, 3, "BER < 7•10-6 after 8 000 000 bits or BER < 10-5 after 160 000 000 bits");
            Report.SetCellProperty(Txsheet, 11, 2, 11, 2, 10, true, 6, -4119);
            Report.SetCellValue(Txsheet, 13, 2, "Channel");
            Report.SetCellValue(Txsheet, 13, 3, "BER 8 000 000");
            Report.SetCellValue(Txsheet, 13, 4, "RX level(dBm)");
            Report.SetCellValue(Txsheet, 13, 6, "BER 8 000 000");
            Report.SetCellValue(Txsheet, 13, 7, "RX level(dBm)");
            Report.SetCellProperty(Txsheet, 13, 2, 13, 4, 10, true, 6, -4119);
            Report.SetCellProperty(Txsheet, 13, 6, 13, 7, 10, true, 6, -4119);
            //Report.SetAutoFit(Txsheet, 2, 2, 13, 7);
            row = 13;
            col = 0;
            TempI = 0;
            Printf(Vi, "CONFigure:BLUetooth:SIGN:OPMode RFT;; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:BTYPe EDR; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:HWINterface1 NONE; *WAI\n");
            //Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:BRATe DH1; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PCONtrol:STEP:ACTion MAX; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PATTern:EDRate PRBS9; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:HOPPing OFF; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:WHITening ON; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:DTX OFF; *WAI\n");
            //Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:DTX:MODE:EDRate SPEC; *WAI\n");
            Printf(Vi, "CONFigure:BLUetooth:SIGN:RXQuality:REPetition SINGleshot; *WAI\n");

            Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:LEVel -20; *WAI\n");
            foreach (string varType in listTEMP)
            {
                Output("Packets type:" + varType);
                col += 3;
                row = 13;
                Printf(Vi, "CONFigure:BLUetooth:SIGN:RXQuality:PACKets " + PacketsNUM[TempI].ToString() + "; *WAI\n");
                Printf(Vi, "CONFigure:BLUetooth:SIGN:CONNection:PACKets:PTYPe:EDRate " + varType + "; *WAI\n");
                TempI++;
                foreach (int varCH in listchannel)
                {
                    Output("Channel:" + varCH.ToString());
                    row += 1;
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "FETCh:BLUetooth:SIGN:CONNection:STATe?\n");
                    Scanf(Vi, "%t", Feedback);
                    if (Feedback.ToString() != "TCON\n")
                    {
                        if (!GetVi()){return;}
                        if (!ConnectCMW()){ CMWinfo.Text = "找不到测试设备，请检查连接设置！";CMWinfo.BackColor = Color.Red; return;}
                    }
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:CHANnel:LOOPback 0," + varCH.ToString() + "; *WAI\n");
                    Report.SetCellValue(Txsheet, row, 2, varCH.ToString());
                    Feedback.Remove(0, Feedback.Length);
                    Printf(Vi, "READ:BLUetooth:SIGN:RXQuality:BER?\n");
                    Failcount = 0;
                    do
                    {
                        Failcount += 1;
                        Feedback.Remove(0, Feedback.Length);
                        Thread.Sleep(500);
                        Scanf(Vi, "%t", Feedback);
                    } while (Feedback.Length == 0 && Failcount <= 80);
                    s = Feedback.ToString().Split(',');
                    Report.SetCellValue(Txsheet, row, col + 1, "-20");
                    Report.SetCellValue(Txsheet, row, col, s[1], "0.0000000");
                }
            }

            //Printf(Vi, "; *WAI\n");

            Report.SetCellProperty(Txsheet, 14, 2, row, 4, 9, false, 4, 1);
            Report.SetCellProperty(Txsheet, 14, 6, row, 7, 9, false, 4, 1);

        }
        public void InitiaSetting()
        {
            Output("Initialization...");
            if (BR.Checked)
            {
                listBurtype.Add("BR");
                if (T1DH1.Checked)
                {
                    listBRtype.Add("DH1");
                }
                if (T1DH3.Checked)
                {
                    listBRtype.Add("DH3");
                }
                if (T1DH5.Checked)
                {
                    listBRtype.Add("DH5");
                }
            }
            if (EDR.Checked)
            {
                listBurtype.Add("EDR");
                if (T2DH1.Checked)
                {
                    listEDRtype2.Add("E21P");
                }
                if (T2DH3.Checked)
                {
                    listEDRtype2.Add("E23P");
                }
                if (T2DH5.Checked)
                {
                    listEDRtype2.Add("E25P");
                }
                if (T3DH1.Checked)
                {
                    listEDRtype3.Add("E31P");
                }
                if (T3DH3.Checked)
                {
                    listEDRtype3.Add("E33P");
                }
                if (T3DH5.Checked)
                {
                    listEDRtype3.Add("E35P");
                }
            }
        }
        public void Output(string log)
        {
            ResultLog.AppendText(DateTime.Now.ToString("HH:mm:ss") +"  "+ log + "\r\n");
        }
        public void Scanf(int vi, string readFmt, StringBuilder arg)
        {
            Visa32.viScanf(vi, readFmt, arg);
            ResultLog.AppendText(arg.ToString() + "\r\n");

        }
        public void Printf(int vi, string writeFmt)
        {
            Visa32.viPrintf(vi, writeFmt);
            EventLog.AppendText(DateTime.Now.ToString("HH:mm:ss") +"  "+ writeFmt + "\r\n");
        }

        private void StartTest_Click(object sender, EventArgs e)
        {
            EventLog.Clear();
            ResultLog.Clear();
            Output("Start test!");
            switch (NunCh.Value)
            {
                case 1:
                    listchannel.Add(0);
                    listchannel.Add(39);
                    listchannel.Add(78);
                    break;
                case 0:
                    if (Channellist.Text == "请输入要测试的频段" || Channellist.Text == "")
                    {
                        CMWinfo.Text = "请先设置测试频段！";
                        CMWinfo.BackColor = Color.Red;
                        return;
                    }
                    else
                    {
                        s = Channellist.Text.Split(',');
                        for (i = 0; i < s.Length; i++)
                        {
                            try { listchannel.Add(int.Parse(s[i])); }
                            catch
                            {
                                CMWinfo.Text = "设置频段格式有误，请重新设置";
                                CMWinfo.BackColor = Color.Red;
                                return;
                            }

                        }
                    }

                    break;
                case 2:
                    for (i = 0; i < 79; i++)
                    {
                        listchannel.Add(i);
                    }


                    break;
            }
            InitiaSetting();
            Report.Create();
            Output("Save report.");
            if (!Report.SaveAs(SaveExcel()))
            {
                Output("No save address selected!");
                Output("End");
                return;
            }
            Output("Save report sucess.");
            if (listBurtype.Count() == 0)
            {
                CMWinfo.Text = "未选择测试项目！";
                CMWinfo.BackColor = Color.Red;
                return;
            }
            if (!GetVi())
            {
                Report.Close();
                return;
            }
            if (!ConnectCMW())
            {
                CMWinfo.Text = "找不到测试设备，请检查连接设置！";
                CMWinfo.BackColor = Color.Red; Report.Close();
                return;
            }
            if (TRMoutputPower.Checked)
            {
                Sett.SelectedTab = Transmitter;
                TRMoutputPower.BackColor = Color.Yellow;
                try
                {
                    GetOutputPower();
                    Output("Test Output Power Success");
                    TRMoutputPower.BackColor = Color.Green;
                }
                catch (Exception EX)
                {
                    Output("Test Output Power Failure");
                    TRMoutputPower.BackColor = Color.Red;
                    Output("Failure log:" + EX.ToString());
                }
                Report.Save();
            }
            if (TRMpowerControl.Checked)
            {
                Sett.SelectedTab = Transmitter;
                TRMpowerControl.BackColor = Color.Yellow;
                try
                {
                    GetPowerControl();
                    Output("Test Power Control Success");
                    TRMpowerControl.BackColor = Color.Green;
                }
                catch (Exception EX)
                {
                    Output("Test Power Control Failure");
                    TRMpowerControl.BackColor = Color.Red;
                    Output("Failure log:" + EX.ToString());
                }
                Report.Save();
            }
            if (TRMfrequenyRange.Checked)
            {
                Sett.SelectedTab = Transmitter;
                TRMfrequenyRange.BackColor = Color.Yellow;
                try
                {
                    GetFrequanceRange();
                    Output("Test Frequance Range Success");
                    TRMfrequenyRange.BackColor = Color.Green;
                }
                catch (Exception EX)
                {
                    Output("Test Frequance Range Failure");
                    TRMfrequenyRange.BackColor = Color.Red;
                    Output("Failure log:" + EX.ToString());
                }
                Report.Save();
            }
            if (TRM20dBBandwidth.Checked)
            {
                Sett.SelectedTab = Transmitter;
                TRM20dBBandwidth.BackColor = Color.Yellow;
                try
                {
                    Get20dBBandwidth();
                    Output("Test 20dB Bandwidth Success");
                    TRM20dBBandwidth.BackColor = Color.Green;
                }
                catch (Exception EX)
                {
                    Output("Test 20dB Bandwidth Failure");
                    Output("Failure log:" + EX.ToString());
                    TRM20dBBandwidth.BackColor = Color.Red;
                }
                Report.Save();
            }
            if (TRMadjacentChannel.Checked)
            {
                Sett.SelectedTab = Transmitter;
                TRMadjacentChannel.BackColor = Color.Yellow;
                try
                {
                    GetAdjacentChannelPower();
                    Output("Test Adjacent Channel Power Success");
                    TRMadjacentChannel.BackColor = Color.Green;
                }
                catch (Exception EX)
                {
                    Printf(Vi, "CONFigure:BLUetooth:MEAS:MEValuation:SACP:BRATe:MEASurement:MODE CH21; *WAI\n");
                    Output("Test Adjacent Channel Power Failure");
                    TRMadjacentChannel.BackColor = Color.Red;
                    Output("Failure log:" + EX.ToString());
                }
                Report.Save();
            }
            if (TRMmodulationCharacteristics.Checked)
            {
                Sett.SelectedTab = Transmitter;
                TRMmodulationCharacteristics.BackColor = Color.Yellow;
                try
                {
                    GetModulationCharacterristics();
                    Output("Test Modulation Characterristics Success");
                    TRMmodulationCharacteristics.BackColor = Color.Green;
                }
                catch (Exception EX)
                {
                    Output("Test Modulation Characterristics Failure");
                    TRMmodulationCharacteristics.BackColor = Color.Red;
                    Output("Failure log:" + EX.ToString());
                }
                Report.Save();
            }
            if (TRMinitialCarrier.Checked)
            {
                Sett.SelectedTab = Transmitter;
                TRMinitialCarrier.BackColor = Color.Yellow;
                try
                {
                    GetCarrierFrequencyTolerance();
                    Output("Test Carrier Frequency Tolerance Success");
                    TRMinitialCarrier.BackColor = Color.Green;
                }
                catch (Exception EX)
                {
                    Output("Test Carrier Frequency Tolerance Failure");
                    TRMinitialCarrier.BackColor = Color.Red;
                    Output("Failure log:" + EX.ToString());
                }
                Report.Save();
            }
            if (TRMcarrierFrequency.Checked)
            {
                Sett.SelectedTab = Transmitter;
                TRMcarrierFrequency.BackColor = Color.Yellow;
                try
                {
                    GetCarrierFrequancyDirft();
                    Output("Test Carrier Frequency Dirft Success");
                    TRMcarrierFrequency.BackColor = Color.Green;
                }
                catch (Exception EX)
                {
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:RFSettings:HOPPing OFF; *WAI\n");
                    Output("Test Carrier Frequency Dirft Failure");
                    TRMcarrierFrequency.BackColor = Color.Red;
                    Output("Failure log:" + EX.ToString());
                }
                Report.Save();
            }
            if (MinRelativeTransmit.Checked)
            {
                Sett.SelectedTab = Transmitter;
                MinRelativeTransmit.BackColor = Color.Yellow;
                try
                {
                    GetEDRrelativeTransmitPower();
                    Output("Test EDR relative Transmit Power Success");
                    MinRelativeTransmit.BackColor = Color.Green;
                }
                catch (Exception EX)
                {
                    Output("Test EDR relative Transmit Power Failure");
                    MinRelativeTransmit.BackColor = Color.Red;
                    Output("Failure log:" + EX.ToString());
                }
                Report.Save();
            }
            if (EDRCarrierFrequency.Checked)
            {
                Sett.SelectedTab = Transmitter;
                EDRCarrierFrequency.BackColor = Color.Yellow;
                try
                {
                    GetEDRcarrirFrequencyStabilityAndModulationAccurcy();
                    Output("Test EDR carrir Frequency Stability And Modulation Accurcy Success");
                    EDRCarrierFrequency.BackColor = Color.Green;

                }
                catch (Exception EX)
                {
                    Output("Test EDR carrir Frequency Stability And Modulation Accurcy Failure");
                    EDRCarrierFrequency.BackColor = Color.Red;
                    Output("Failure log:" + EX.ToString());
                }
                Report.Save();
            }
            if (DifferentialPhaseEncoding.Checked)
            {
                Sett.SelectedTab = Transmitter;
                DifferentialPhaseEncoding.BackColor = Color.Yellow;
                try
                {
                    GetEDRdifferentialPhaseEncoding();
                    Output("Test EDR differential Phase Encoding Success");
                    DifferentialPhaseEncoding.BackColor = Color.Green;
                }
                catch (Exception EX)
                {
                    Printf(Vi, "CONFigure:BLUetooth:SIGN:TMODe LOOPback; *WAI\n");
                    Output("Test EDR differential Phase Encoding Failure");
                    DifferentialPhaseEncoding.BackColor = Color.Red;
                    Output("Failure log:" + EX.ToString());
                }
                Report.Save();
            }
            if (RCVsensitivityMeasurements.Checked)
            {
                Sett.SelectedTab = Receiver;
                RCVsensitivityMeasurements.BackColor = Color.Yellow;
                try
                {
                    GetSensitivityMultiSlot();
                    Output("Test Sensitivity Multi-Slot Success");
                    RCVsensitivityMeasurements.BackColor = Color.Green;
                }
                catch (Exception EX)
                {
                    Output("Test Sensitivity Multi-Slot Failure");
                    RCVsensitivityMeasurements.BackColor = Color.Red;
                    Output("Failure log:" + EX.ToString());
                }
                Report.Save();
            }
            if (RCVMaximumInputLevel.Checked)
            {
                Sett.SelectedTab = Receiver;
                RCVMaximumInputLevel.BackColor = Color.Yellow;
                try
                {
                    GetMaximumInputLevel();
                    Output("Test Maximum Input Level Success");
                    RCVMaximumInputLevel.BackColor = Color.Green;
                }
                catch (Exception EX)
                {
                    Output("Test Maximum Input Level Failure");
                    RCVMaximumInputLevel.BackColor = Color.Red;
                    Output("Failure log:" + EX.ToString());
                }
                Report.Save();
            }
            if (EDRsensitivity.Checked)
            {
                Sett.SelectedTab = Receiver;
                EDRsensitivity.BackColor = Color.Yellow;
                try
                {
                    GetEDRSensitivity();
                    Output("Test EDR Sensitivity Success");
                    EDRsensitivity.BackColor = Color.Green;
                }
                catch (Exception EX)
                {
                    Output("Test EDR Sensitivity Failure");
                    EDRsensitivity.BackColor = Color.Red;
                    Output("Failure log:" + EX.ToString());
                }
                Report.Save();
            }
            if (EDRBERfloorPerformance.Checked)
            {
                Sett.SelectedTab = Receiver;
                EDRBERfloorPerformance.BackColor = Color.Yellow;
                try
                {
                    GetEDRBERFloorPerformance();
                    Output("Test EDR BER Floor Performance Success");
                    EDRBERfloorPerformance.BackColor = Color.Green;
                }
                catch (Exception EX)
                {
                    Output("Test EDR BER Floor Performance Failure");
                    EDRBERfloorPerformance.BackColor = Color.Red;
                    Output("Failure log:" + EX.ToString());
                }
                Report.Save();
            }
            if (EDRmaximunInputLevel.Checked)
            {
                Sett.SelectedTab = Receiver;
                EDRmaximunInputLevel.BackColor = Color.Yellow;
                try
                {
                    GetEDRMaximumInputLevel();
                    Output("Test EDR Maximum Input Level Success");
                    EDRmaximunInputLevel.BackColor = Color.Green;
                }
                catch (Exception EX)
                {
                    Output("Test EDR Maximum Input Level Failure");
                    EDRmaximunInputLevel.BackColor = Color.Red;
                    Output("Failure log:" + EX.ToString());
                }
                Report.Save();
            }           
            Report.Save();
            Report.Close();
            Visa32.viClose(Vi);
            Output("Test complete!");
            return;
        }
        private void NunCh_Scroll(object sender, EventArgs e)
        {
            switch (NunCh.Value)
            {
                case 1:
                    Channellist.Text = "0,39,78";
                    Channellist.Enabled = false;
                    break;
                case 0:
                    Channellist.Enabled = true;
                    Channellist.Text = "请输入要测试的频段";

                    break;
                case 2:
                    Channellist.Text = "All channels";
                    Channellist.Enabled = false;
                    break;
            }
        }
        private string SaveExcel()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                InitialDirectory = Path,
                //打开的文件选择对话框上的标题  
                Title = "请保存报告",
                //设置文件类型  
                Filter = "Excel(*.xlsx)|*.xlsx",
                //设置默认文件类型显示顺序  
                FilterIndex = 1,
                //保存对话框是否记忆上次打开的目录  
                RestoreDirectory = true
            };
            //按下确定选择的按钮  
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {

                //获得文件路径  
                string localFilePath = saveFileDialog.FileName.ToString();
                //获取文件路径，不带文件名  
                //FilePath = localFilePath.Substring(0, localFilePath.LastIndexOf("\\"));  
                //获取文件名，带后缀名，不带路径  
                string fileNameWithSuffix = localFilePath.Substring(localFilePath.LastIndexOf("\\") + 1);
                return localFilePath;
            }
            else
            {
                CMWinfo.Text = "请保存测试结果！";
                CMWinfo.BackColor = Color.Red;
                return "";

            }

        }
    }
    internal sealed class Visa32
    {
        // --------------------------------------------------------------------------------
        //  Adapted from visa32.bas which was distributed by VXIplug&play Systems Alliance
        //  Distributed By Agilent Technologies, Inc.
        // --------------------------------------------------------------------------------
        // -------------------------------------------------------------------------
        public const int VI_SPEC_VERSION = 4194304;
        #region - Resource Template Functions and Operations ----------------------------
        [DllImportAttribute("VISA32.DLL", EntryPoint = "#141", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true)]
        public static extern int viOpenDefaultRM(out int sesn);
        [DllImportAttribute("VISA32.DLL", EntryPoint = "#128", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true)]
        public static extern int viGetDefaultRM(out int sesn);
        [DllImportAttribute("VISA32.DLL", EntryPoint = "#131", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true)]
        public static extern int viOpen(int sesn, string viDesc, int mode, int timeout, out int vi);
        [DllImportAttribute("VISA32.DLL", EntryPoint = "#132", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true)]
        public static extern int viClose(int vi);
        [DllImportAttribute("VISA32.DLL", EntryPoint = "#269", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true, CallingConvention = CallingConvention.Cdecl)]
        public static extern int viPrintf(int vi, string writeFmt);
        [DllImportAttribute("VISA32.DLL", EntryPoint = "#271", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true, CallingConvention = CallingConvention.Cdecl)]
        public static extern int viScanf(int vi, string readFmt, StringBuilder arg);
        #endregion
    }
}





namespace Microsoft.Office.Interop.ExcelEdit
{
    /// <SUMMARY>
    /// Microsoft.Office.Interop.ExcelEdit 的摘要说明
    /// </SUMMARY>
    public class ExcelEdit
    {
        public string mFilename;
        public Excel.Application app;
        public Excel.Workbooks wbs;
        public Excel.Workbook wb;
        public Excel.Worksheets wss;
        public Excel.Worksheet ws;
        public Excel.Range Allrange;
        public ExcelEdit()
        {
            //
            // TODO: 在此处添加构造函数逻辑
            //
        }
        public void Create()//创建一个Microsoft.Office.Interop.Excel对象
        {
            app = new Excel.Application
            {
                StandardFont = "Consolas"
            };
            wbs = app.Workbooks;
            wb = wbs.Add(true);
        }
        public void Open(string FileName)//打开一个Microsoft.Office.Interop.Excel文件
        {
            app = new Excel.Application();
            wbs = app.Workbooks;
            wb = wbs.Add(FileName);
            //wb = wbs.Open(FileName, 0, true, 5,"", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "t", false, false, 0, true,Type.Missing,Type.Missing);
            //wb = wbs.Open(FileName,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing);
            mFilename = FileName;
        }
        public Excel.Worksheet GetSheet(string SheetName)
        //获取一个工作表
        {
            Excel.Worksheet s = (Excel.Worksheet)wb.Worksheets[SheetName];
            return s;
        }
        public Excel.Worksheet AddSheet(string SheetName)
        //添加一个工作表
        {
            Excel.Worksheet Tempsheet = (Excel.Worksheet)wb.Sheets.get_Item(wb.Sheets.Count);
            Excel.Worksheet sheet = (Excel.Worksheet)wb.Worksheets.Add(Type.Missing, Tempsheet, Type.Missing, Type.Missing);
            sheet.Name = SheetName;
            //Allrange = sheet.Columns;
            //Allrange.WrapText = false;
            return sheet;
        }

        public void DelSheet(string SheetName)//删除一个工作表
        {
            ((Excel.Worksheet)wb.Worksheets[SheetName]).Delete();
        }
        /*
        public Excel.Worksheet ReNameSheet(string OldSheetName, string NewSheetName)//重命名一个工作表一
        {
            Excel.Worksheet s = (Excel.Worksheet)wb.Worksheets[OldSheetName];
            s.Name = NewSheetName;
            return s;
        }
        */
    
        public Excel.Worksheet ReNameSheet(Excel.Worksheet Sheet, string NewSheetName)//重命名一个工作表二
        {

            Sheet.Name = NewSheetName;

            return Sheet;
        }

        public void SetCellValue(Excel.Worksheet ws, int x, int y, object value)
        //ws：要设值的工作表     X行Y列     value   值
        {
            ws.Cells[x, y] = value;
            ws.Range[ws.Cells[x, y], ws.Cells[x, y]].HorizontalAlignment = Excel.Constants.xlLeft;
        }
        public void SetCellValue(string ws, int x, int y, object value)
        //ws：要设值的工作表的名称 X行Y列 value 值
        {

            GetSheet(ws).Cells[x, y] = value;
           // GetSheet(ws).Cells[x, y].HorizontalAlignment = Excel.Constants.xlLeft;
           // GetSheet(ws).Cells[x, y].WrapText = false;
        }
        public void SetCellValue(Excel.Worksheet ws, int x, int y, object value,string Fomat)
        //ws：要设值的工作表的名称 X行Y列 value 值
        {

            ws.Cells[x, y] = value;
            ws.Range[ws.Cells[x, y], ws.Cells[x, y]].HorizontalAlignment = Excel.Constants.xlLeft;
            ws.Range[ws.Cells[x, y], ws.Cells[x, y]].WrapText = false;
            ws.Range[ws.Cells[x, y], ws.Cells[x, y]].NumberFormatLocal = Fomat;
        }
        public void SetAutoFit(Excel.Worksheet ws, int Startx, int Starty, int Endx, int Endy)
        {
            ws.Range[ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]].WrapText = false;

        }
        public void SetFomules(Excel.Worksheet ws, int Startx, int Starty, int Endx, int Endy,string Fomulas)
        {
            ws.Range[ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]].NumberFormatLocal = Fomulas;
            ws.Range[ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]].EntireColumn.AutoFit();

        }

        public void SetCellProperty(Excel.Worksheet ws, int Startx, int Starty, int Endx, int Endy, int size, bool bold, int color, int linestyle)
        //设置一个单元格的属性   字体，   大小，颜色   ，对齐方式
        {
            //size = 12;
            //color = Excel.Constants.xlAutomatic;
            // ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Name = "宋体";
            ws.Range[ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]].Font.Bold = bold;
            ws.Range[ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]].Font.Size = size;
            //ws.Range[ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]].Font.Color = color;
            ws.Range[ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]].Interior.ColorIndex = color;
            //实线1，双线-4119，颜色3红4绿6黄
            ws.Range[ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]].Borders.LineStyle = linestyle;
            ws.Range[ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]].HorizontalAlignment = Excel.Constants.xlLeft;
        }
        /*
        public void SetCellProperty(string wsn, int Startx, int Starty, int Endx, int Endy, int size, string name, Excel.Constants color, Excel.Constants HorizontalAlignment)
        {
            //name = "宋体";
            //size = 12;
            //color = Microsoft.Office.Interop.Excel.Constants.xlAutomatic;
            //HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;

            Excel.Worksheet ws = GetSheet(wsn);
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Interior.ColorIndex = 20;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Size = size;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Color = color;

            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).HorizontalAlignment = HorizontalAlignment;
        }
        */


        public void UniteCells(Excel.Worksheet ws, int x1, int y1, int x2, int y2)
        //合并单元格
        {
            ws.get_Range(ws.Cells[x1, y1], ws.Cells[x2, y2]).Merge(Type.Missing);
        }

        public void UniteCells(string ws, int x1, int y1, int x2, int y2)
        //合并单元格
        {
            GetSheet(ws).get_Range(GetSheet(ws).Cells[x1, y1], GetSheet(ws).Cells[x2, y2]).Merge(Type.Missing);

        }


        public void InsertTable(System.Data.DataTable dt, string ws, int startX, int startY)
        //将内存中数据表格插入到Microsoft.Office.Interop.Excel指定工作表的指定位置 为在使用模板时控制格式时使用一
        {

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {
                    GetSheet(ws).Cells[startX + i, j + startY] = dt.Rows[i][j].ToString();

                }

            }

        }
        public void InsertTable(System.Data.DataTable dt, Excel.Worksheet ws, int startX, int startY)
        //将内存中数据表格插入到Microsoft.Office.Interop.Excel指定工作表的指定位置二
        {

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {

                    ws.Cells[startX + i, j + startY] = dt.Rows[i][j];

                }

            }

        }


        public void AddTable(System.Data.DataTable dt, string ws, int startX, int startY)
        //将内存中数据表格添加到Microsoft.Office.Interop.Excel指定工作表的指定位置一
        {

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {

                    GetSheet(ws).Cells[i + startX, j + startY] = dt.Rows[i][j];

                }

            }

        }
        public void AddTable(System.Data.DataTable dt, Excel.Worksheet ws, int startX, int startY)
        //将内存中数据表格添加到Microsoft.Office.Interop.Excel指定工作表的指定位置二
        {


            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {

                    ws.Cells[i + startX, j + startY] = dt.Rows[i][j];

                }
            }

        }
        public void InsertPictures(string Filename, string ws)
        //插入图片操作一
        {
            GetSheet(ws).Shapes.AddPicture(Filename, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, 10, 150, 150);
            //后面的数字表示位置
        }

        //public void InsertPictures(string Filename, string ws, int Height, int Width)
        //插入图片操作二
        //{
        //    GetSheet(ws).Shapes.AddPicture(Filename, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, 10, 150, 150);
        //    GetSheet(ws).Shapes.get_Range(Type.Missing).Height = Height;
        //    GetSheet(ws).Shapes.get_Range(Type.Missing).Width = Width;
        //}
        //public void InsertPictures(string Filename, string ws, int left, int top, int Height, int Width)
        //插入图片操作三
        //{

        //    GetSheet(ws).Shapes.AddPicture(Filename, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, 10, 150, 150);
        //    GetSheet(ws).Shapes.get_Range(Type.Missing).IncrementLeft(left);
        //    GetSheet(ws).Shapes.get_Range(Type.Missing).IncrementTop(top);
        //    GetSheet(ws).Shapes.get_Range(Type.Missing).Height = Height;
        //    GetSheet(ws).Shapes.get_Range(Type.Missing).Width = Width;
        //}

        public void InsertActiveChart(Excel.XlChartType ChartType, string ws, int DataSourcesX1, int DataSourcesY1, int DataSourcesX2, int DataSourcesY2, Microsoft.Office.Interop.Excel.XlRowCol ChartDataType)
        //插入图表操作
        {
            ChartDataType = Excel.XlRowCol.xlColumns;
            wb.Charts.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            {
                wb.ActiveChart.ChartType = ChartType;
                wb.ActiveChart.SetSourceData(GetSheet(ws).get_Range(GetSheet(ws).Cells[DataSourcesX1, DataSourcesY1], GetSheet(ws).Cells[DataSourcesX2, DataSourcesY2]), ChartDataType);
                wb.ActiveChart.Location(Excel.XlChartLocation.xlLocationAsObject, ws);
            }
        }
        public bool Save()
        //保存文档
        {
            if (mFilename == "")
            {
                return false;
            }
            else
            {
                try
                {
                    wb.Save();
                    return true;
                }

                catch (Exception ex)
                {
                    return false;
                }
            }
        }
        public bool SaveAs(object FileName)
        //文档另存为
        {
            try
            {
                wb.SaveAs(FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                return true;

            }

            catch (Exception ex)
            {
                return false;

            }
        }
        public void Close()
        //关闭一个Microsoft.Office.Interop.Excel对象，销毁对象
        {
            //wb.Save();
            wb.Close(Type.Missing, Type.Missing, Type.Missing);
            wbs.Close();
            app.Quit();
            wb = null;
            wbs = null;
            app = null;
            GC.Collect();
        }
    }
}