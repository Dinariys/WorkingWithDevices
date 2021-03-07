// На основе документации: https://zipstore.ru/auxpage_cmd-tiger-p-tiger-pro/

using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Windows.Forms;

namespace DeviceManager
{

    public class MettlerToledoTiger
    {
        private TcpClient tcpClient;
        private NetworkStream stream;
        private IPEndPoint RemoteIpEndPoint;
        /// <summary>
        /// Справочник товаров
        /// </summary>
        private readonly BindingSource bsDictionary;
        /// <summary>
        /// Свойства товаров
        /// </summary>
        private readonly BindingSource bsProperties;

        public MettlerToledoTiger(BindingSource bsDictionary, BindingSource bsProperties, string Host, int Port = 3001)
        {
            this.bsDictionary = bsDictionary;
            this.bsProperties = bsProperties;
            var HostPort = Host.Split(':');
            Host = HostPort[0];
            Port = HostPort.Length == 2 ? Convert.ToInt32(HostPort[1]) : Port;
            RemoteIpEndPoint = new IPEndPoint(IPAddress.Parse(Host), Port);
        }

        public bool LoadGoodsScale()
        {
            try
            {
                tcpClient = new TcpClient();
                tcpClient.Connect(RemoteIpEndPoint);

                var CommandBody = new List<byte>();
                var CommandPacket = new List<byte>();
                var CommandFull = new List<byte>();
                short PageCount = 0;
                foreach (DataRowView row in bsDictionary)
                {
                    if (!CmdBody(CommandBody, row))
                        continue;

                    PageCount++;
                }

                if (PageCount == 0)
                {
                    AppLogs.MessageLog("Mettler Toledo. Нет данных для выгрузки", ELogStatus.Warning, true);
                    return false;
                }

                GetPacketHeader(CommandPacket, CmdHeader.Count, CommandBody.Count, PageCount);

                CommandFull.AddRange(CommandPacket.ToArray());
                CommandFull.AddRange(CmdHeader.ToArray());
                CommandFull.AddRange(CommandBody.ToArray());
                var CRC16 = Calc(CommandFull.ToArray());
                CommandFull.AddRange(BitConverter.GetBytes(CRC16).Reverse());
                CommandFull.Insert(0, 0x02); // Стартовый байт

                //AppLogs.MessageLog("CommandPacket:" + ByteArrayToString(CommandPacket.ToArray()));
                //AppLogs.MessageLog($"CmdHeader [{CmdHeader.Count}]:" + ByteArrayToString(CmdHeader.ToArray()));
                //AppLogs.MessageLog($"CommandBody [{CommandBody.Count}]:" + ByteArrayToString(CommandBody.ToArray()));

                //AppLogs.MessageLog("CommandFull:" + ByteArrayToString(CommandFull.ToArray()));

                SendData(CommandFull);
                Disconnect();
                return true;
            }
            catch (Exception Ex)
            {
                //AppLogs.MessageLog(Ex, "Метод [LoadGoodScale]");
                Disconnect();
                return false;
            }
        }

        private bool CmdBody(List<byte> CommandBody, DataRowView good)
        {
            if (!int.TryParse(good["Code"].ToString(), out int GoodCode))
                return false;

            // Ограничение на поле (PLU, PLU2) 30 символов
            string TempName = good["PrintName"].ToString().Length > 30
                ? good["PrintName"].ToString().Substring(0, 30)
                : good["PrintName"].ToString().PadRight(30, ' ');

            string TempName2 = string.Empty;
            if (good["PrintName"].ToString().Length > 60)
            {
                TempName2 = good["PrintName"].ToString().Substring(30, 30);
            }
            else if (good["PrintName"].ToString().Length > 30 && good["PrintName"].ToString().Length < 60)
            {
                TempName2 = good["PrintName"].ToString().Substring(30).PadRight(30, ' ');
            }
            else
            {
                TempName2 = TempName2.PadRight(30, ' ');
            }

            int TempPrice = Convert.ToInt32(Convert.ToDecimal(good["PriceOut"].ToString()) * 100);

            short DateValidity = 0;
            bsProperties.Filter = string.Format("GoodID = {0} AND TypeID = 16", Convert.ToInt32(good["ID"]));
            if (bsProperties.Count > 0)
            {
                short.TryParse(((DataRowView)bsProperties.Current)["Value"].ToString(), out DateValidity);
            }

            byte[] PLUNo = BitConverter.GetBytes(GoodCode);                         // L06
            byte[] ArcticleNo = Encoding.Default.GetBytes($"{GoodCode:D13}");
            byte[] PLUName = Encoding.GetEncoding("CP866").GetBytes(TempName);
            byte[] PLUName2 = Encoding.GetEncoding("CP866").GetBytes(TempName2);
            byte[] Empty = Encoding.GetEncoding("CP866").GetBytes(" ");
            // Цена
            byte[] UnitPrice = BitConverter.GetBytes(TempPrice);                    // L08
            byte TaxRate = 0;
            byte Tare = 0;
            byte[] FixWeight = { 0x00, 0x00, 0x00, 0x00 };                          // L11
            // Группа всегда 1
            byte[] GroupNo = { 0x01, 0x00 };                                        // S04
            // Весовой или штучный товар
            byte[] FlagF04 = good["Barcode"].ToString().Contains('.') 
                ? new byte[] { 0x00, 0x00 } 
                : new byte[] { 0x01, 0x00 };                                        // F04
            byte[] BestByDateOffset = BitConverter.GetBytes(DateValidity);          // S03
            byte[] SellByDateOffset = { 0x00, 0x00 };                               // S03
            byte[] ExtraTxtNumber = { 0x00, 0x00 };                                 // S03

            CommandBody.AddRange(PLUNo);
            CommandBody.AddRange(ArcticleNo);
            CommandBody.AddRange(PLUName);
            CommandBody.AddRange(PLUName2);
            CommandBody.Add(Empty[0]);
            CommandBody.AddRange(UnitPrice);
            CommandBody.Add(TaxRate);
            CommandBody.AddRange(FixWeight);
            CommandBody.Add(Tare);
            CommandBody.AddRange(new byte[] { 0x00, 0x00 });                        // S04 - Nothing
            CommandBody.AddRange(GroupNo);
            CommandBody.AddRange(FlagF04);
            CommandBody.AddRange(BestByDateOffset);
            CommandBody.AddRange(SellByDateOffset);
            CommandBody.AddRange(ExtraTxtNumber);
            return true;
        }

        private void GetPacketHeader(List<byte> CommandPacket, int LengthCmdHeader, int LengthCmdBody, short PageCount)
        {
            //short PageLength = Convert.ToInt16(LengthCmdBody);
            short PageLength = 100;
            short TotalLength = Convert.ToInt16(LengthCmdHeader + PageCount * PageLength);

            CommandPacket.AddRange(BitConverter.GetBytes(TotalLength));
            CommandPacket.AddRange(BitConverter.GetBytes(PageCount));
            CommandPacket.AddRange(BitConverter.GetBytes(PageLength));
        }

        private List<byte> CmdHeader
        {
            get
            {
                short CommanSend = 207;
                byte CommandResponse = 0x00;                        // [1] байт на отправку или ответ
                byte[] Command = BitConverter.GetBytes(CommanSend); // [2] байта на комнду
                byte Control = 0x00;                                // [1] байт на тип команды
                short DepartNo = 1;                                 // [2] байта на номер департамена 
                short DeviceNo = 0;                                 // [2] байта на номер устройства
                var listData = new List<byte>();
                listData.Add(CommandResponse);
                listData.AddRange(Command);
                listData.Add(Control);
                listData.AddRange(BitConverter.GetBytes(DepartNo).Reverse());
                listData.AddRange(BitConverter.GetBytes(DeviceNo));
                return listData;
            }
        }

        private void SendData(List<byte> listData)
        {
            try
            {
                stream = null;
                stream = tcpClient.GetStream();
                stream.Write(listData.ToArray(), 0, listData.ToArray().Length);
            }
            catch (Exception Ex)
            {
                //AppLogs.MessageLog(Ex, "Метод [SendData: Scale]");
            }
        }

        private void Disconnect()
        {            
            stream?.Close();
            tcpClient.Close();
        }

        /// <summary>
        /// Расчёт CRC
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private static ushort Calc(byte[] data)
        {
            ushort wCRC = 0;
            for (int i = 0; i < data.Length; i++)
            {
                wCRC ^= (ushort)(data[i] << 8);
                for (int j = 0; j < 8; j++)
                {
                    if ((wCRC & 0x8000) != 0)
                        wCRC = (ushort)((wCRC << 1) ^ 0x1021);
                    else
                        wCRC <<= 1;
                }
            }
            return wCRC;
        }

        private string ByteArrayToString(byte[] ba)
        {
            StringBuilder hex = new StringBuilder(ba.Length * 2);
            foreach (byte b in ba)
                hex.AppendFormat("{0:x2} ", b);
            return hex.ToString();
        }
    }
}
