using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;

namespace MiDaysCommerce.DeviceManager
{
    public class Feliks : IDisposable
    {
        private string COMNumber { get; set; }
        private int BaudRate { get; set; }
        private int TimeOut { get; set; }
        /// <summary>
        /// Таймаут для возврата к исходному значению
        /// </summary>
        private int TempTimeout { get; set; }
        private bool IsOpenMsgBox = false;

        /// <summary>
        /// Your serial port
        /// </summary>
        private SerialPort serialPort;
        private AutoResetEvent receiveNow;

        #region Константы
        /// <summary>
        /// [ENQ] Запрос
        /// </summary>
        private const byte Request = 0x05;
        /// <summary>
        /// [ACK] Подтверждение
        /// </summary>
        private const byte Confirm = 0x06;
        /// <summary>
        /// Начало текста
        /// </summary>
        private const byte STX = 0x02;
        /// <summary>
        /// Конец текста
        /// </summary>
        private const byte ETX = 0x03;
        /// <summary>
        /// Конец передачи
        /// </summary>
        private const byte EOT = 0x04;
        /// <summary>
        /// Экранирование управляющих символов
        /// </summary>
        private const byte DLE = 0x10;
        private const byte EMPTY = 0x00;
        #endregion

        /// <summary>
        /// Пароль по умолчанию
        /// </summary>
        private byte DefaultPassword = 0x30;
        private byte[] response = null;

        /// <summary>
        /// Шаблон чека
        /// </summary>
        private TemplateReceipt Template;
        /// <summary>
        /// Набор данных
        /// </summary>
        private Scripts.DataTemplate data;

        public Feliks(TemplateReceipt Template, Scripts.DataTemplate data, string COMNumber, int BaudRate, int TimeOut)
        {
            this.Template = Template;
            this.data = data;
            this.COMNumber = string.IsNullOrWhiteSpace(COMNumber) ? "COM1" : COMNumber;
            this.BaudRate = BaudRate;
            this.TimeOut = TimeOut;
            TempTimeout = TimeOut;
        }

        #region Работа с COM портом

        /// <summary>
        /// Указание настроек для подключения к COM
        /// </summary>
        private void SetPort()
        {
            serialPort = new SerialPort(COMNumber, BaudRate)
            {
                Parity = Parity.None,

                Handshake = Handshake.None,
                DataBits = 8,
                StopBits = StopBits.One,
                RtsEnable = true,
                DtrEnable = true,
                WriteTimeout = TimeOut,
                ReadTimeout = TimeOut
            };
        }

        /// <summary>
        /// Открыть COM порт
        /// </summary>
        /// <returns>При успешном открытии вернет true</returns>
        private bool Open()
        {
            try
            {
                if (serialPort == null)
                    return false;

                if (serialPort.IsOpen)
                {
                    if (!Close())
                    {
                        //AppLogs.MessageLog("Не удалось закрыть порт", ELogStatus.Warning);
                        return false;
                    }
                }

                receiveNow = new AutoResetEvent(false);
                serialPort.Open();
                serialPort.DataReceived += new SerialDataReceivedEventHandler(serialPort_DataReceived);
                //AppLogs.MessageLog("Успешное открытие порта");
                return true;
            }
            catch (Exception Ex)
            {
                //AppLogs.MessageLog(Ex, "Не удалось открыть порт", ELogStatus.Critical, false, false);
                return false;
            }
        }

        /// <summary>
        /// Закрываем COM порт
        /// </summary>
        /// <returns>Если успешно вернут true</returns>
        private bool Close()
        {
            try
            {
                if (serialPort == null)
                    return false;

                if (!serialPort.IsOpen)
                    return false;

                serialPort.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void serialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                if (e.EventType == SerialData.Chars)
                {
                    receiveNow.Set();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private byte[] ExecuteCommand(byte[] cmd, int countByte)
        {
            try
            {
                serialPort.DiscardOutBuffer();
                serialPort.DiscardInBuffer();
                receiveNow.Reset();
                serialPort.Write(cmd, 0, cmd.Length);
                serialPort.ReadTimeout = TimeOut;

                byte[] input = ReadResponse(countByte);
                return input;
            }
            catch
            {
                return null;
            }
        }

        private byte[] ReadResponse(int byteCount)
        {
            byte[] buffer = new byte[byteCount];
            try
            {
                List<byte> listByte = new List<byte>();
                while (listByte.Count < byteCount)
                {
                    byte b = (byte)serialPort.ReadByte();
                    listByte.Add(b);
                }
                buffer = listByte.ToArray();

                return buffer;
            }
            catch (Exception Ex)
            {
                //AppLogs.MessageLog(Ex, string.Empty, ELogStatus.Critical, false, false);
                return buffer;
            }
        }

        #endregion

        /// <summary>
        /// Установить соединение
        /// </summary>
        /// <returns></returns>
        private bool OpenConnection()
        {
            SetPort();
            bool IsResult = Open();
            //if(!IsResult)
            //    data.Response = AppLogs.MessageLog("Феликс. Нет связи с фискальным регистратором.", ELogStatus.Warning, IsOpenMsgBox);

            return IsResult;
        }
                
        /// <summary>
        /// Разорвать соединение
        /// </summary>
        private bool CloseConnection()
        {
            return Close();
        }

        private bool CheckConnection()
        {
            response = ExecuteCommand(new byte[] { Request }, 1);
            if (response == null)
            {
                //data.Response = AppLogs.MessageLog("Феликс. Нет связи с устройством. Нет ответа в отведенный промежуток времени.", ELogStatus.Warning, IsOpenMsgBox);
                CloseConnection();
                return false;
            }

            if (response[0] != Confirm)
            {
                //data.Response = AppLogs.MessageLog("Феликс. Нет связи с устройством. Нет подтверждения в отведенный промежуток времени.");
                CloseConnection();
                return false;
            }

            return true;
        }

        private bool CommandClientAddress()
        {
            if (string.IsNullOrWhiteSpace(data.ClientAddress))
                return true;

            byte Mode = EMPTY;
            if (data.ClientAddress.Contains("@"))
                Mode = 0x02;
            else
            {
                // Первым символом в номере обязательно должен быть символ «+»
                data.ClientAddress = data.ClientAddress.Contains("+") ? data.ClientAddress : "+" + data.ClientAddress;
            }

            // Звонок на номер телефона
            const byte SMS = 0x70;
            var TempCommand = new List<byte>();
            TempCommand.AddRange(new byte[] { EMPTY, EMPTY, SMS });
            TempCommand.Add(Mode);
            TempCommand.AddRange(Encoding.GetEncoding(866).GetBytes(data.ClientAddress));

            var command = GetPreparedCommand(TempCommand.ToArray());
            // Отправляем основную команду и ожидаем подтверждения
            response = ExecuteCommand(command, 1);

            if (!IsFromDeviceRead())
                return false;

            if (!CheckError(response))
                return false;

            return IsToDeviceWrite();
        }

        private bool CommandOperatorName()
        {
            if (string.IsNullOrWhiteSpace(data.User))
                return true;

            // Звонок на номер телефона
            const byte USER = 0x7E;
            var TempCommand = new List<byte>();
            TempCommand.AddRange(new byte[] { EMPTY, EMPTY, USER, 0x19 });
            TempCommand.AddRange(Encoding.GetEncoding(866).GetBytes(data.User));

            var command = GetPreparedCommand(TempCommand.ToArray());
            // Отправляем основную команду и ожидаем подтверждения
            response = ExecuteCommand(command, 1);

            if (!IsFromDeviceRead())
                return false;

            if (!CheckError(response))
                return false;

            return IsToDeviceWrite();
        }

        private bool CommandOperatorCode()
        {
            if (string.IsNullOrWhiteSpace(data.UserUCode))
                return true;

            // Звонок на номер телефона
            const byte USER = 0x7E;
            var TempCommand = new List<byte>();
            TempCommand.AddRange(new byte[] { EMPTY, EMPTY, USER, 0x1C });
            TempCommand.AddRange(Encoding.GetEncoding(866).GetBytes(data.UserUCode));

            var command = GetPreparedCommand(TempCommand.ToArray());
            // Отправляем основную команду и ожидаем подтверждения
            response = ExecuteCommand(command, 1);

            if (!IsFromDeviceRead())
                return false;

            if (!CheckError(response))
                return false;

            return IsToDeviceWrite();
        }

        /// <summary>
        /// Получить текущее состояние ККТ
        /// </summary>
        /// <param name="Mode"></param>
        /// <returns></returns>
        private int GetCurrentCode(int Mode)
        {
            // Запрос кода состояния ККТ
            const byte E = 0x45;
            var command = GetPreparedCommand(new byte[] { EMPTY, EMPTY, E });
            // Отправляем основную команду и ожидаем подтверждения
            response = ExecuteCommand(command, 1);

            if (!IsFromDeviceRead())
                return 0;

            GetPreparedResponse(response);
            if (!CheckError(response))
                return 0;
            
            if (response[2] == Mode)
                return Mode;

            if (response[2] == 0)
                return response[2];

            CommandExitMode();
            return 0;
        }

        /// <summary>
        /// Комманда выхода из текущего режима
        /// </summary>
        private void CommandExitMode()
        {
            IsToDeviceWrite();

            // Выход из текущего режима
            const byte R = 0x52;

            var command = GetPreparedCommand(new byte[] { EMPTY, EMPTY, R });
            // Отправляем основную команду и ожидаем подтверждения
            response = ExecuteCommand(command, 1);

            IsFromDeviceRead();
        }

        /// <summary>
        /// Комманда на смену режима работы
        /// </summary>
        /// <param name="Mode">Номер режима</param>
        /// <returns>В случае успешной смены вернёт true, иначе вернёт false</returns>
        private bool CommandChangeMode(int Mode)
        {
            // Вход в режим
            const byte V = 0x56;
            // Устанавливаемые режимы (двоично-десятичное число):
            // 1 - режим регистрации;
            // 2 - оборота по текущей смена без закрытия;
            // 3 - оборота по текущей смена с закрытием;
            // 4 - режим программирования;
            // 5 - режим налогового инспектора;

            if (GetCurrentCode(Mode) == Mode)
                return true;
            
            if (!IsToDeviceWrite())
                return false;

            var command = GetPreparedCommand(new byte[] { EMPTY, EMPTY, V, (byte)Mode, EMPTY, EMPTY, EMPTY, DefaultPassword });
            // Отправляем основную команду и ожидаем подтверждения
            response = ExecuteCommand(command, 1);

            if (!IsFromDeviceRead())
                return false;

            return CheckError(response);
        }

        /// <summary>
        /// От устройства
        /// </summary>
        private bool IsFromDeviceRead()
        {
            if (response == null)
                return false;

            // Проверяем подтвердило ли устройство получение
            if (response[0] != Confirm)
                return false;

            // Завершаем передачу
            response = ExecuteCommand(new byte[] { EOT }, 1);
            // Проверяем есть ли запрос от устройства
            if (response[0] != Request)
                return false;

            // Подтверждаем получение запроса и получаем новые данные
            response = ExecuteCommand(new byte[] { Confirm }, 6);
            return true;
        }

        /// <summary>
        /// К устройству
        /// </summary>
        private bool IsToDeviceWrite()
        {
            // Отправляем подтверждение о получении
            response = ExecuteCommand(new byte[] { Confirm }, 1);
            // Проверяем завершена ли передача с устройством
            if (response[0] != EOT)
                return false;
            // Отправляем запрос
            response = ExecuteCommand(new byte[] { Request }, 1);
            // Проверяем подтвердило ли устройство получение
            return response[0] == Confirm;
        }

        /// <summary>
        /// Получить подготовленную комманду
        /// </summary>
        /// <param name="DataWrite"></param>
        /// <returns></returns>
        private byte[] GetPreparedCommand(byte[] DataWrite)
        {
            var command = new List<byte>();
            for (int i = 0; i < DataWrite.Length; i++)
            {
                // Экранируем специальные команды
                if (DataWrite[i] == ETX)
                    command.AddRange(new byte[] { DLE, DataWrite[i] });
                else if (DataWrite[i] == DLE)
                    command.AddRange(new byte[] { DLE, DLE });
                else
                    command.Add(DataWrite[i]);
            }
            command.Add(ETX);

            var crc = CalcCRC(command.ToArray());
            command.Insert(0, STX);
            command.Insert(command.Count, crc);
            
            return command.ToArray();
        }

        /// <summary>
        /// Получить подготовленный ответ
        /// </summary>
        /// <param name="DataRead"></param>
        private void GetPreparedResponse(byte[] DataRead)
        {
            var tempResponse = new List<byte>();
            for (int i = 0; i < DataRead.Length; i++)
            {
                if (DataRead[i] != DLE)
                {
                    tempResponse.Add(DataRead[i]);
                    continue;
                }
                int index = i + 1 > DataRead.Length ? i : i + 1;
                if (DataRead[index] == ETX)
                    continue;
                
                tempResponse.Add(DataRead[i]);
            }

            response = tempResponse.ToArray();
        }

        /// <summary>
        /// Команда снятия X отчёта
        /// </summary>
        /// <returns></returns>
        private bool CommandReportX()
        {
            if (!CheckConnection())
                return false;

            if (!CommandChangeMode(2))
                return false;

            if (!IsToDeviceWrite())
                return false;

            CommandOperatorName();
            CommandOperatorCode();

            // Оборот по текущей смене без закрытия
            const byte g = 0x67;
            // Тип Отчёта: 1 – Оборот по текущей смена без закрытия
            const int TypeReport = 1;
            var command = GetPreparedCommand(new byte[] { EMPTY, EMPTY, g, TypeReport });
            // Отправляем основную команду и ожидаем подтверждения
            response = ExecuteCommand(command, 1);

            if (!IsFromDeviceRead())
                return false;

            if (!CheckError(response))
                return false;

            return IsToDeviceWrite();
        }

        /// <summary>
        /// Открыть смену
        /// </summary>
        private bool CommandOpenShift()
        {
            if (!CheckConnection())
                return false;

            if (!CommandChangeMode(1))
                return false;

            if (!IsToDeviceWrite())
                return false;

            CommandOperatorName();
            CommandOperatorCode();

            // Открыть смену
            const byte SHIFT = 0x9A;
            var command = GetPreparedCommand(new byte[] { EMPTY, EMPTY, SHIFT });
            // Отправляем основную команду и ожидаем подтверждения
            response = ExecuteCommand(command, 1);

            if (!IsFromDeviceRead())
                return false;

            if (!CheckError(response))
                return false;

            return IsToDeviceWrite();
        }

        private bool CommandOpenCheck(int TypeCheck)
        {
            CommandOperatorName();
            CommandOperatorCode();

            // Открытие чека
            const byte CHECK = 0x92;
            var command = GetPreparedCommand(new byte[] { EMPTY, EMPTY, CHECK, EMPTY, (byte)TypeCheck });
            // Отправляем основную команду и ожидаем подтверждения
            response = ExecuteCommand(command, 1);

            if (!IsFromDeviceRead())
                return false;

            if (!CheckError(response))
                return false;

            return IsToDeviceWrite();
        }

        /// <summary>
        /// Разбить строку на определенные разряды
        /// </summary>
        /// <param name="sr"></param>
        /// <param name="size"></param>
        /// <param name="fixedSize"></param>
        /// <remarks>https://ru.stackoverflow.com/a/478202/190858</remarks>
        IEnumerable<string> Split(TextReader sr, int size, bool fixedSize = true)
        {
            while (sr.Peek() >= 0)
            {
                var buffer = new char[size];
                var c = sr.ReadBlock(buffer, 0, size);
                yield return fixedSize ? new String(buffer) : new String(buffer, 0, c);
            }
        }

        IEnumerable<string> Split(string s, int size, bool fixedSize = true)
        {
            var sr = new StringReader(s);
            return Split(sr, size, fixedSize);
        }

        /// <summary>
        /// Получить число
        /// </summary>
        private byte[] GetDigit(int Data)
        {
            string TempValue = string.Format("{0:D10}", Data);
            var result = Split(TempValue, 2).Select(s => Convert.ToByte(s, 16)).ToArray();
            return result;
        }

        /// <summary>
        /// Комманда печати наименования товара (услуги)
        /// </summary>
        private bool CommandPrintText(string TextPrint, byte Mode = 0x4C)
        {
            // Печатаемые символы (X до 96 символов) - кодовой странице 866 MSDOS. 
            TextPrint = TextPrint.Length > 96 ? TextPrint.Substring(0, 96) : TextPrint;

            // Открытие чека
            byte TEXT = Mode;
            var TempCommand = new List<byte>();
            TempCommand.AddRange(new byte[] { EMPTY, EMPTY, TEXT });
            TempCommand.AddRange(Encoding.GetEncoding(866).GetBytes(TextPrint));
            var command = GetPreparedCommand(TempCommand.ToArray());
            // Отправляем основную команду и ожидаем подтверждения
            response = ExecuteCommand(command, 1);

            if (!IsFromDeviceRead())
                return false;

            if (!CheckError(response))
                return false;

            return IsToDeviceWrite();
        }

        /// <summary>
        /// Команда регистрации позиции прихода/расхода
        /// </summary>
        private bool CommandRegistrationPosition(DataRow good, int OperType)
        {
            byte[] Price = GetDigit(Convert.ToInt32(data.PriceOut * 100));
            byte[] Quantity = GetDigit(Convert.ToInt32(data.Quantity * 1000));
            byte Department = EMPTY;
            byte Discount = 0x01;
            byte[] VatValue = { 0x00, 0x00, 0x00 };
            byte[] DiscountSum = { 0x00, 0x00, 0x00, 0x00, 0x00 };
            byte VatGroupCode = GetVatGroup(Convert.ToInt32(good["VatGroupCode"]));
            byte[] VatSum = { 0x00, 0x00, 0x00, 0x00, 0x00 };
            byte[] Sum = GetDigit(Convert.ToInt32(Math.Round(data.PriceOut * data.Quantity, 2, MidpointRounding.AwayFromZero) * 100));
            byte Sing = (byte)Convert.ToInt32(good["TypeGoodFPID"]);
            byte SignPay = 0x04;

            // Регистрация прихода/расхода

            byte REGISTRATION = 0x48;
            if (OperType == 2)
                REGISTRATION = 0x57;

            var TempCommand = new List<byte>();
            TempCommand.AddRange(new byte[] { EMPTY, EMPTY, REGISTRATION, EMPTY });
            TempCommand.AddRange(Price);
            TempCommand.AddRange(Quantity);
            TempCommand.Add(Department);
            TempCommand.Add(Discount);
            TempCommand.AddRange(VatValue);
            TempCommand.AddRange(DiscountSum);
            TempCommand.Add(VatGroupCode);
            TempCommand.AddRange(VatSum);
            TempCommand.AddRange(Sum);
            TempCommand.Add(Sing);
            TempCommand.Add(SignPay);

            var command = GetPreparedCommand(TempCommand.ToArray());
            // Отправляем основную команду и ожидаем подтверждения
            response = ExecuteCommand(command, 1);

            if (!IsFromDeviceRead())
                return false;

            if (!CheckError(response))
                return false;

            return IsToDeviceWrite();
        }

        /// <summary>
        /// Команда оплаты
        /// </summary>
        private bool CommandPay(int Type, int Sum)
        {
            // Тип оплаты
            byte PayType = GetPayType(Type);
            byte[] PaySum = GetDigit(Sum);

            // Расчёт по чеку
            const byte PAY = 0x99;
            var TempCommand = new List<byte>();
            TempCommand.AddRange(new byte[] { EMPTY, EMPTY, PAY, EMPTY });
            TempCommand.Add(PayType);
            TempCommand.AddRange(PaySum);

            var command = GetPreparedCommand(TempCommand.ToArray());
            // Отправляем основную команду и ожидаем подтверждения
            response = ExecuteCommand(command, 1);

            if (!IsFromDeviceRead())
                return false;

            if (!CheckError(response))
                return false;

            return IsToDeviceWrite();
        }

        /// <summary>
        /// Команда внесения/выплаты денег
        /// </summary>
        private bool CommandMoneyInOut(int Sum)
        {
            if (!CheckConnection())
                return false;

            if (!CommandChangeMode(1))
                return false;

            if (!IsToDeviceWrite())
                return false;

            CommandOperatorName();
            CommandOperatorCode();

            // Внесение денег
            const byte MONEYIN = 0x49;
            // Выплата денег
            const byte MONEYOUT = 0x4F;

            byte CurrentMode = MONEYIN;
            if (Sum < 0)
            {
                CurrentMode = MONEYOUT;
                Sum = Math.Abs(Sum);
            }

            byte[] CurrentSum = GetDigit(Sum);

            var TempCommand = new List<byte>();
            TempCommand.AddRange(new byte[] { EMPTY, EMPTY, CurrentMode, EMPTY });
            TempCommand.AddRange(CurrentSum);

            var command = GetPreparedCommand(TempCommand.ToArray());
            // Отправляем основную команду и ожидаем подтверждения
            response = ExecuteCommand(command, 1);

            if (!IsFromDeviceRead())
                return false;

            if (!CheckError(response))
                return false;

            return IsToDeviceWrite();
        }

        /// <summary>
        /// Команда закрытия чека
        /// </summary>
        private bool CommandCloseCheck()
        {
            // Закрытие чека
            const byte CLOSE = 0x4A;
            var TempCommand = new List<byte>();
            TempCommand.AddRange(new byte[] { EMPTY, EMPTY, CLOSE, EMPTY, EMPTY });

            var command = GetPreparedCommand(TempCommand.ToArray());
            // Отправляем основную команду и ожидаем подтверждения
            response = ExecuteCommand(command, 1);

            if (!IsFromDeviceRead())
                return false;

            if (!CheckError(response))
                return false;

            return IsToDeviceWrite();
        }

        /// <summary>
        /// Комманда аннулирования чека
        /// </summary>
        /// <returns></returns>
        private bool CommandCancelCheck()
        {
            // Аннулирование чека
            const byte CANCEL = 0x59;
            var command = GetPreparedCommand(new byte[] { EMPTY, EMPTY, CANCEL });
            // Отправляем основную команду и ожидаем подтверждения
            response = ExecuteCommand(command, 1);

            if (!IsFromDeviceRead())
                return false;

            if (!CheckError(response))
                return false;

            //AppLogs.MessageLog("Феликс выполнено аннулирование ранее открытого чека.");
            return true;
        }

        private bool CommandReportZ()
        {
            if (!CheckConnection())
                return false;

            if (!CommandChangeMode(3))
                return false;

            if (!IsToDeviceWrite())
                return false;

            CommandOperatorName();
            CommandOperatorCode();

            // Оборот по текущей смене с закрытием
            const byte Z = 0x5A;
            var command = GetPreparedCommand(new byte[] { EMPTY, EMPTY, Z });
            // Отправляем основную команду и ожидаем подтверждения
            response = ExecuteCommand(command, 1);

            if (!IsFromDeviceRead())
                return false;

            if (!CheckError(response))
                return false;

            return IsToDeviceWrite(); ;
        }

        public bool Sale(EPrintTypeCode TypePrint)
        {
            if (TypePrint == EPrintTypeCode.Service)
                return CommonService();

            return Common(1);
        }

        public bool RefundSale(EPrintTypeCode TypePrint)
        {
            if (TypePrint == EPrintTypeCode.Service)
                return CommonService();

            return Common(2);
        }

        /// <summary>
        /// Общий метод
        /// </summary>
        /// <param name="OperType">Тип операции</param>
        /// <returns>В случае успеха вернёт true, иначе false</returns>
        private bool Common(int OperType)
        {
            try
            {
                if (!OpenConnection())
                    return false;

                if (!CommandOpenShift())
                {
                    CloseConnection();
                    return false;
                }

                if (!CommandOpenCheck(OperType))
                {
                    CloseConnection();
                    return false;
                }

                CommandClientAddress();

                decimal docSum = 0;
                if (data.dsDetails != null)
                {
                    //AppLogs.MessageLog("Передачи товарных позиций в чек");
                    foreach (DataRow good in data.dsDetails.Tables["Details"].Select(data.PrintGroups))
                    {
                        // Подсчет суммы для не печатаемых товаров
                        if (string.IsNullOrWhiteSpace(good["NotPrint"].ToString()) == false && (bool)good["NotPrint"] == true)
                        {
                            docSum = docSum + (Convert.ToDecimal(good["PriceOut"]) * Convert.ToDecimal(good["Quantity"]));
                            continue;
                        }

                        methodsFiscal.FillDataRow(good, ref data);
                        if (!CommandPrintText(data.GoodName))
                            return false;

                        if (!CommandRegistrationPosition(good, OperType))
                            return false;
                    }
                }

                foreach (DataRow row in data.dtPayClientSum.Rows)
                {
                    if (Convert.ToInt32(row["ID"]) != (int)ECodeTypePay.CreditCart)
                        continue;

                    if (row["Sum"] == null || Convert.ToDecimal(row["Sum"]) == 0)
                        continue;

                    int PayType = Convert.ToInt32(row["ID"]);
                    int PaySum = Convert.ToInt32((Convert.ToDecimal(row["Sum"]) - docSum) * 100);

                    if (!CommandPay(PayType, PaySum))
                        return false;
                }

                foreach (DataRow row in data.dtPayClientSum.Rows)
                {
                    if (Convert.ToInt32(row["ID"]) == (int)ECodeTypePay.CreditCart)
                        continue;

                    if (row["Sum"] == null || Convert.ToDecimal(row["Sum"]) == 0)
                        continue;

                    int PayType = Convert.ToInt32(row["ID"]);
                    int PaySum = Convert.ToInt32((Convert.ToDecimal(row["Sum"]) - docSum) * 100);

                    if (!CommandPay(PayType, PaySum))
                        return false;
                }

                if (!CommandCloseCheck())
                    return false;

                CloseConnection();
                //AppLogs.MessageLog("Феликс. Фискальный чек №0 распечатан успешно. Документ закрыт.");
                return true;
            }
            catch
            {
                CloseConnection();
                //data.Response = "Феликс. Возникла не опознанная ошибка. Пожалуйста обратитесь к специалисту.";
                return false;
            }
        }

        /// <summary>
        /// Общий метод для не фискального чека
        /// </summary>
        /// <returns>В случае успеха вернёт true, иначе false</returns>
        private bool CommonService()
        {
            try
            {
                if (!OpenConnection())
                    return false;

                if (!CheckConnection())
                    return false;

                if (!CommandChangeMode(1))
                    return false;

                if (!IsToDeviceWrite())
                    return false;

                PrintString(EPartTemplate.Header);

                if (data.dsDetails != null)
                {
                    //AppLogs.MessageLog("Передачи товарных позиций в чек");
                    foreach (DataRow good in data.dsDetails.Tables["Details"].Select(data.PrintGroups))
                    {
                        PrintString(EPartTemplate.Details);
                    }
                }

                PrintString(EPartTemplate.Footer);

                CloseConnection();
                //AppLogs.MessageLog("Феликс. Служебный чек №0 распечатан успешно. Документ закрыт.");
                return true;
            }
            catch
            {
                CloseConnection();
                //data.Response = "Феликс. Возникла не опознанная ошибка. Пожалуйста обратитесь к специалисту.";
                return false;
            }
        }

        public bool ReportX()
        {
            TimeOut = 20000;
            if (!OpenConnection())
                return false;

            if (!CommandReportX())
            {
                CloseConnection();
                return false;
            }

            CloseConnection();
            //AppLogs.MessageLog("Феликс. X отчет снят успешно.");
            TimeOut = TempTimeout;
            return true;
        }

        public bool ReportZ(bool IsAuto = false)
        {
            TimeOut = 20000;
            if (!OpenConnection())
                return false;

            if (!CommandReportZ())
            {
                CloseConnection();
                return false;
            }

            //if (IsAuto)

            CloseConnection();
            //AppLogs.MessageLog("Феликс. Z отчет снят успешно.");
            TimeOut = TempTimeout;
            return true;
        }

        /// <summary>
        /// Внесение или выплата
        /// </summary>
        /// <returns>В случае успеха вернёт true, иначе false</returns>
        public bool MoneyInOut()
        {
            if (!OpenConnection())
                return false;

            int CurrentSum = Convert.ToInt32(data.SumMoneyInOut * 100);
            if (!CommandMoneyInOut(CurrentSum))
            {
                //data.Response = AppLogs.MessageLog(@"Феликс. Не удалось выполнить операцию внесения\выплаты. Пожалуйста обратитесь к специалисту.", ELogStatus.Warning, IsOpenMsgBox);
                CloseConnection();
                return false;
            }

            CloseConnection();
            //data.Response = AppLogs.MessageLog($"Феликс. Операция внесения/выплата на сумму {data.SumMoneyInOut:0.00} выполнена успешно.", ELogStatus.Information, IsOpenMsgBox);
            return true;
        }

        private bool CheckError(byte[] DataRead)
        {
            if (DataRead.Length < 3)
                return false;

            if (DataRead[0] != STX)
                return false;

            if (DataRead[1] == 0x00)
                return true;

            if (DataRead[1] == 0x55)
            {
                switch (DataRead[2])
                {
                    case 0x00:
                        return true;
                    case 0x01:
                    case 0x02:
                    case 0x03:
                        return true;
                    case 0xE8:
                        //data.Response = AppLogs.MessageLog("Феликс. Смена превысила 24 часа. Снимите Z отчёта и повторите операцию", ELogStatus.Warning, IsOpenMsgBox);
                        // Отправляем подтверждение о получении
                        response = ExecuteCommand(new byte[] { Confirm }, 1);
                        // Проверяем завершена ли передача с устройством
                        if (response[0] != EOT)
                            return false;
                            
                        return false;
                    case 0x6A:
                        //data.Response = AppLogs.MessageLog("Феликс. Неверный тип чека", ELogStatus.Warning, IsOpenMsgBox);
                        CloseConnection();
                        return false;
                    case 0x67:
                        //data.Response = AppLogs.MessageLog("Феликс. Отсутствует бумага.", ELogStatus.Warning, IsOpenMsgBox);
                        return false;
                    case 0x76:
                        //data.Response = AppLogs.MessageLog("Феликс. Смена закрыта операция не возможна.", ELogStatus.Warning, IsOpenMsgBox);
                        CloseConnection();
                        return false;
                    case 0x8C:
                        //data.Response = AppLogs.MessageLog("Феликс. Устройство вернуло не опознанную ошибку. Пожалуйста свяжитесь со специалистом!", ELogStatus.Warning, IsOpenMsgBox);
                        //AppLogs.MessageLog("Не верный пароль кассира", ELogStatus.Warning);
                        CloseConnection();
                        return false;
                    case 0x9A:
                        //data.Response = AppLogs.MessageLog("Феликс. Устройство вернуло не опознанную ошибку. Пожалуйста свяжитесь со специалистом!", ELogStatus.Warning, IsOpenMsgBox);
                        //AppLogs.MessageLog("Чек закрыт операция невозможна", ELogStatus.Warning);
                        CloseConnection();
                        return false;
                    case 0x9C:
                        //AppLogs.MessageLog("Смена уже открыта");
                        return true;
                    case 0x9B:
                        //AppLogs.MessageLog("Феликс. Есть открытый чек, будет предпринята попытка аннулирования.", ELogStatus.Warning);
                        if (!IsToDeviceWrite())
                        {
                            //data.Response = AppLogs.MessageLog("Феликс. Не удалось аннулировать ранее открытый чек.", ELogStatus.Warning, IsOpenMsgBox);
                            CloseConnection();
                            return false;
                        }

                        return CommandCancelCheck();
                    case 0x66:
                        //data.Response = AppLogs.MessageLog("Феликс. Устройство вернуло не опознанную ошибку. Пожалуйста свяжитесь со специалистом!", ELogStatus.Warning, IsOpenMsgBox);
                        //AppLogs.MessageLog("Команда не реализуется в данном режиме ККТ", ELogStatus.Warning);
                        CloseConnection();
                        return false;

                    default:
                        //data.Response = AppLogs.MessageLog("Феликс. Устройство вернуло не опознанную ошибку. Пожалуйста свяжитесь со специалистом!", ELogStatus.Warning, IsOpenMsgBox);
                        CloseConnection();
                        return false;
                }
            }
            return false;
        }

        /// <summary>
        /// Печать строки
        /// </summary>
        /// <param name="ePart">Часть чека</param>
        private void PrintString(EPartTemplate ePart)
        {
            string Text;

            if (Template == null)
                return;

            switch (ePart)
            {
                case EPartTemplate.Header:
                    foreach (var s in Template.Header.ToArray())
                    {
                        Text = s.Line;
                        CommandPrintText(Scripts.GetDataRow(Text, data), 0x5E);
                        Text = string.Empty;
                    }
                    break;
                case EPartTemplate.Details:
                    foreach (var s in Template.Detail.ToArray())
                    {
                        Text = s.Line;
                        CommandPrintText(Scripts.GetDataRow(Text, data), 0x5E);
                        Text = string.Empty;
                    }
                    break;
                case EPartTemplate.Footer:
                    foreach (var s in Template.Footer.ToArray())
                    {
                        Text = s.Line;

                        if (s.Line.Contains("[PayTypeSum]"))
                        {
                            decimal TempSum = 0;
                            foreach (DataRow row in data.dtPayClientSum.Rows)
                            {
                                if (row["Sum"] != null && Convert.ToDouble(row["Sum"]) > 0)
                                {
                                    TempSum += Convert.ToDecimal(row["Sum"]);
                                    CommandPrintText($"{row["Name"]}: {row["Sum"]:0.00}", 0x5E);
                                }
                            }

                            if (TempSum == 0)
                            {
                                CommandPrintText("БЕЗ ОПЛАТЫ", 0x5E);
                            }
                        }
                        else
                        {
                            CommandPrintText(Scripts.GetDataRow(Text, data), 0x5E);                           
                        }
                        Text = string.Empty;
                    }
                    break;
            }
        }

        /// <summary>
        /// Расчёт контрольного числа
        /// </summary>
        private byte CalcCRC(byte[] command)
        {
            int result = 0;
            for (int i = 0; i < command.Count(); i++)
            {
                result ^= command[i];
            }

            return Convert.ToByte(result);
        }

        /// <summary>
        /// Получить тип оплаты для устройства
        /// </summary>
        /// <param name="PayType">Тип оплаты</param>
        public byte GetPayType(int PayType)
        {
            switch ((ECodeTypePay)PayType)
            {
                case ECodeTypePay.Cash:
                    return 0x01;
                case ECodeTypePay.CreditCart:
                    return 0x02;

                default:
                    return 0x01;
            }
        }

        /// <summary>
        /// Получить код группы НДС для ФР
        /// </summary>
        private byte GetVatGroup(int VatGroupCode)
        {
            switch (VatGroupCode)
            {
                // Ставка НДС 20%
                case 1:
                    return 0x01;
                // Ставка НДС 10%
                case 2:
                    return 0x02;
                // Ставка НДС 0%
                case 3:
                    return 0x05;
                // НДС не облагается
                case 4:
                    return 0x06;
                // Ставка НДС расчётная 20/120
                case 5:
                    return 0x03;
                // Ставка НДС расчётная 10/110
                case 6:
                    return 0x04;

                // НДС не облагается
                default:
                    return 0x06;
            }
        }

        #region Dispose
        private bool disposed = false;

        // реализация интерфейса IDisposable.
        public void Dispose()
        {
            Dispose(true);
            // подавляем финализацию
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    // Освобождаем управляемые ресурсы
                }
                // освобождаем неуправляемые объекты
                disposed = true;
            }
        }

        // Деструктор
        ~Feliks()
        {
            Dispose(false);
        }
        #endregion
    }
}
