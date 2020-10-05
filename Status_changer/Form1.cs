using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using teemtalk;
using System.Globalization;
using System.Diagnostics;
using NLog;
using Excel = Microsoft.Office.Interop.Excel;


namespace Status_changer
{
    
    public partial class Form1 : Form
    {

        private static Logger logger = LogManager.GetCurrentClassLogger(); // Nlog

        public Form1()
        {
            InitializeComponent();
        }



        static teemtalk. Application teemApp;

        public string EventDepot { get; private set; }

        bool Stop = false;

        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                

                if (textBox_login.Text == "")
                {
                    MessageBox.Show("Вы не ввели логин", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if (textBox_pw.Text == "")
                {
                    MessageBox.Show("Вы не ввели пароль", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }





                //Login into Mainframe
                var login = textBox_login.Text;
                var password = textBox_pw.Text;

                //Login for auto testing
                //login = "Q583eyj";
                //password = "IamGr00t";



                teemApp = new teemtalk.Application();

                teemApp.CurrentSession.Name = "Mainframe";

                teemApp.CurrentSession.Network.Protocol = ttNetworkProtocol.ProtocolWinsock;
                teemApp.CurrentSession.Network.Hostname = "mainframe.gb.tntpost.com";
                teemApp.CurrentSession.Network.Telnet.Port = 23;
                teemApp.CurrentSession.Network.Telnet.Name = "IBM-3278-2-E";
                teemApp.CurrentSession.Emulation = ttEmulations.IBM3270Emul;

                teemApp.CurrentSession.Network.Connect();

                teemApp.Visible = Properties.Settings.Default.isVisible;


                var host = teemApp.CurrentSession.Host;
                var disp = teemApp.CurrentSession.Display;


                teemApp.CurrentSession.Keyboard.Macros.Add("<VK_RETURN>", "<VK_SEPARATOR>", true);



                ForAwait(35, 16, "INTERNATIONAL");

                host.Send("SM");
                host.Send("<ENTER>");

                ForAwait(13, 23, "USER ID");
                Thread.Sleep(2000);
                host.Send(login);
                host.Send("<TAB>");
                host.Send(password);
                host.Send("<ENTER>");

                Thread.Sleep(2000);
                if (teemApp.CurrentSession.Display.CursorCol == 40)
                {
                    TeemTalkClose();
                    MessageBox.Show("Вы ввели неверный логин или пароль. Введите правильные данные и нажмите кнопку START", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;

                }
                else if (teemApp.CurrentSession.Display.CursorCol == 35)
                {
                    TeemTalkClose();
                    MessageBox.Show("Ваш пароль устарел. Измените пароль в Mainframe, введите правильные данные и нажмите кнопку START", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;

                }

                // Создаем Client Data txt файл
                string ClientDataPath = @"ClientData";
                if (!Directory.Exists(ClientDataPath)) //Если папки нет...
                    Directory.CreateDirectory(ClientDataPath); //...создадим ее
                string ClientDataName = "ClientData_" + DateTime.Now.ToString("ddMMMyy_HHmm", CultureInfo.GetCultureInfo("en-us")) + ".txt";
                string destClientData = Path.Combine(ClientDataPath, ClientDataName);
                StreamWriter ClientData = new StreamWriter(destClientData, true);
                ClientData.WriteLine("#CS User: " + login);
                ClientData.Close();
                                
                ForAwait(2, 2, "Command");
                host.Send("2");
                host.Send("<ENTER>");

                Thread.Sleep(2000);
                if (teemApp.CurrentSession.Display.CursorCol == 01)
                {
                    TeemTalkClose();
                    MessageBox.Show("Пользователь "+login+" уже авторизован в Terminal I. Выйдете из сессии Terminal I и нажмите кнопку START", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;

                }

                logger.Debug("User:"+login, this.Text); //LOG 

                ForAwait(20, 7, "Job Description");
                host.Send("<F12>");
                Thread.Sleep(500);
                if (disp.CursorRow != 2)


                host.Send("JK04");
                logger.Debug("JK04", this.Text); //LOG
                host.Send("<ENTER>");
                Thread.Sleep(2200);

                if (teemApp.CurrentSession.Display.CursorCol == 25)// Закрытие всплывающего окна, которое нужно закрыть
                {
                    host.Send("<F12>");
                }

                //Проверка на возможность доступа в JK04              
                if (disp.ScreenData[73, 1, 4] != "JK04")
                {
                    TeemTalkClose();
                    MessageBox.Show("Пользователь " + login + " не имеет доступа в  JK04", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string ClientName = "";
                string Tel1 = "";
                string Tel2 = "";

                string wrClientName = "";
                string wrTel1 = "";
                string wrTel2 = "";

                do
                {
                    //btnStop.Enabled = true;

                    if (disp.ScreenData[73, 1, 4] == "JK04" & teemApp.CurrentSession.Display.CursorRow == 11)
                    {                        
                        ClientName = disp.ScreenData[19, 8, 30];
                        Tel1 = disp.ScreenData[19, 9, 7];
                        Tel2 = disp.ScreenData[30, 9, 9];                     


                        if (ClientName != wrClientName || Tel1 != wrTel1 || Tel2 != wrTel2 )
                        {
                            //string Name = disp.ScreenData[19, 8, 10];
                            ClientData = new StreamWriter(destClientData, true);
                            ClientData.Write(ClientName + "; ");
                            ClientData.Close();

                            //string Tel1 = disp.ScreenData[19, 9, 7];
                            ClientData = new StreamWriter(destClientData, true);
                            ClientData.Write(Tel1 + "; ");
                            ClientData.Close();

                            //string Tel2 = disp.ScreenData[30, 9, 9];
                            ClientData = new StreamWriter(destClientData, true);
                            ClientData.WriteLine(Tel2 + ";");
                            ClientData.Close();


                            wrClientName = ClientName;
                            wrTel1 = Tel1;
                            wrTel2 = Tel2;
                        }
                    }
                    
                    //if(Stop == true)
                    //{
                    //    break;
                        
                    //}

                } while (true);

                


                //teemApp.Close();
                //foreach (Process proc in Process.GetProcessesByName("teem2k"))
                //{
                //proc.Kill();
                //}
                //teemApp.Application.Close();
                //Thread.Sleep(1000);
                //host.Send("<ENTER>");

                // Закрываем TeemTalk
                TeemTalkClose();
                logger.Debug("TeemTalkNormalClose", this.Text); //LOG

                btnStart.Enabled = true;

                //Открыть текстовый файл
                Process.Start(destClientData);

                            
                                            

            }
            catch (Exception ex)
            {
                // Вывод сообщения об ошибке
                logger.Debug(ex.ToString());
            }




        }
              

            static void TeemTalkClose()// Закрываем TeemTalk
        {

            teemApp.CurrentSession.Network.Close();
            Thread.Sleep(500);
            teemApp.Close();
        }


        static bool ForAwait(short col, short row, string keyword)
        {
            byte count = 0;
            
                do
                {
                    count++;
                    
                    if (count > 70)
                    {
                        teemApp.CurrentSession.Network.Close();
                        Thread.Sleep(1000);
                        teemApp.Close();

                        System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("teem2k");

                        foreach (System.Diagnostics.Process p in process)
                        {
                            if (!string.IsNullOrEmpty(p.ProcessName))
                            {
                                try
                                {
                                    p.Kill();
                                }

                                catch (Exception ex)
                                {
                                    // Вывод сообщения об ошибке
                                    logger.Debug(ex.ToString());
                                }
                            }
                        }

                        return false;
                    }

                    Thread.Sleep(100);

                } while ((teemApp.CurrentSession.Display.ScreenData[col, row, (short)keyword.Length] != keyword));
            return true;
        }

        static bool ForAwaitRow(short keyword)
        {
            byte count = 0;

            do
            {
                count++;

                if (count > 70)
                {
                    teemApp.CurrentSession.Network.Close();
                    Thread.Sleep(1000);
                    teemApp.Close();

                    System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("teem2k");

                    foreach (System.Diagnostics.Process p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName))
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch (Exception ex)
                            {
                                // Вывод сообщения об ошибке
                                logger.Debug(ex.ToString());
                            }
                        }
                    }

                    return false;
                }

                Thread.Sleep(100);

            } while ((teemApp.CurrentSession.Display.CursorRow != keyword));
            return true;
        }
        static bool ForAwaitCol(short keyword)
        {
            byte count = 0;

            do
            {
                count++;

                if (count > 70)
                {
                    teemApp.CurrentSession.Network.Close();
                    Thread.Sleep(1000);
                    teemApp.Close();

                    System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("teem2k");

                    foreach (System.Diagnostics.Process p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName))
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch (Exception ex)
                            {
                                // Вывод сообщения об ошибке
                                logger.Debug(ex.ToString());
                            }
                        }
                    }

                    return false;
                }

                Thread.Sleep(100);

            } while ((teemApp.CurrentSession.Display.CursorCol != keyword));
            return true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void maskedTextBox1_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
