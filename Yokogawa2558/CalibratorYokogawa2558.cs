using Facility;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO.Ports;
using System.Linq;
using System.Threading;
using Keysight.Visa;

namespace Yokogawa2558
{
    // вопросы по диапазонам частоты
    // вопросы по выключению выхода
    // 
    public class CalibratorYokogawa2558
    {
        [FacilityMetadata(Denotation = "Калибратор", SupportedDevices = "Yokogawa 2558", Interfaces = "GPIB", IsCalibrator = true, Version = 1)]
        public class Yokogawa2558 : Facility.Facility
        {
            #region Members
            private GpibSession session;
            #endregion

            #region Methods
            public override bool Open(byte address, SerialPort serialPort = null)
            {
                session = new ResourceManager().Open($"GPIB{address / 32}::{address % 32}::INSTR") as GpibSession;
                if (session != null)
                {
                    session.Clear();
                    session.TerminationCharacterEnabled = true;
                    _calibrator = new CalibratorYokogawa(session);
                    return true;
                }
                return false;
            }

            public override void Close()
            {
                if (Calibrator != null)
                {
                    session.Dispose();
                    _calibrator = null;
                }
            }
            #endregion
        }

        public class CalibratorYokogawa : Calibrator
        {
            #region Members
            private GpibSession session;
            private string OldRangeCurrent;
            private string OldRangeVoltage;
            private string OldValueFrecuency;
            private bool IsOutputOn;
            #endregion

            #region Construction
            public CalibratorYokogawa(GpibSession session)
            {
                this.session = session;
                Reset();
            }
            #endregion

            #region Methods
           
            /// <summary>
            /// Устанавливает переменное напряжение
            /// </summary>
            /// <param name="value">Среднеквадратичное (RMS) значение напряжения</param>
            /// <returns></returns>
            public override bool SetACVoltage(double value)
            {
                Busy = true;
                //словарь, где ключ - диапазон, а значение - это параметр команды, и множитель
                var HighRange = new Dictionary<string, int[]>() {
                    { "0.1", new int[] { 1, 100000 } },
                    { "1", new int[] { 2, 10000 } },
                    { "10", new int[] { 3, 1000 } },
                    { "100",new int[] { 4, 100 } },
                    { "300",new int[] { 5, 10 } },
                    { "1000", new int[] { 6, 10 } } };

                try
                {
                    if (value > 1000)
                        throw new Exception("Установленное значение напряжения выходит за пределы допустимого диапазона!");
                    
                    
                    var NowRange = HighRange.FirstOrDefault(rang => value <= double.Parse(rang.Key, new CultureInfo("EN")));


                    if (OldRangeVoltage != NowRange.Key)
                    {
                        if (IsOutputOn)
                        {
                            session.FormattedIO.WriteLine("O0");
                            session.AssertTrigger();
                            IsOutputOn = false;
                        }

                        OldRangeVoltage = NowRange.Key;
                    }
                    //устанавливаем значение
                    session.FormattedIO.WriteLine($"V{NowRange.Value[0]}S{(value * NowRange.Value[1]).ToString("00000", new CultureInfo("EN"))}");
                    session.AssertTrigger();

                    if (IsOutputOn.Equals(false))
                    {
                        session.FormattedIO.WriteLine("O1");
                        session.AssertTrigger();
                        IsOutputOn = true;
                    }

                }
                catch (Exception ex)
                {
                    Reset();
                    throw ex;
                }
                Busy = false;
                return true;
            }

            /// <summary>
            /// Устанавливает переменный ток
            /// <param name="value">Среднеквадратичное значение тока</param>
            /// <returns></returns>
            public override bool SetACCurrent(double value)
            {
                Busy = true;
                //словарь, где ключ - диапазон, а значение - это параметр команды, и множитель
                var HighRange = new Dictionary<string, int[]>()
                {
                    { "0.1", new int[] { 1, 100000 } },
                    { "1", new int[] { 2, 10000 } },
                    { "10", new int[] { 3, 1000 } },
                    { "50",new int[] { 4, 100 } } 
                };
                   

                try
                {
                    if (value > 50)
                        throw new Exception("Установленное значение тока выходит за пределы допустимого диапазона!");


                    var NowRange = HighRange.FirstOrDefault(rang => value <= double.Parse(rang.Key, new CultureInfo("EN")));


                    if (OldRangeCurrent != NowRange.Key)
                    {
                        if (IsOutputOn)
                        {
                            session.FormattedIO.WriteLine("O0");
                            session.AssertTrigger();
                            IsOutputOn = false;
                        }

                        OldRangeCurrent = NowRange.Key;
                    }
                    //устанавливаем значение
                    session.FormattedIO.WriteLine($"A{NowRange.Value[0]}S{(value * NowRange.Value[1]).ToString("00000", new CultureInfo("EN"))}");
                    session.AssertTrigger();

                    if (IsOutputOn.Equals(false))
                    {
                        session.FormattedIO.WriteLine("O1");
                        session.AssertTrigger();
                        IsOutputOn = true;
                    }

                }
                catch (Exception ex)
                {
                    Reset();
                    throw ex;
                }
                Busy = false;
                return true;
            }

            /// <summary>
            /// Устанавливает частоту 
            /// </summary>
            /// <param name="value">Значение частоты сигнала</param>
            /// <returns></returns>
            public override bool SetFrequency(double value)
            {
                Busy = true;
                try
                {

                    if (value < 16.0 || value > 6000.0)
                        throw new Exception("Установленное значение частоты выходит за пределы допустимого диапазона!");

                    var val = value.ToString("0.#", new CultureInfo("EN"));
                    if (val != OldValueFrecuency)
                    {
                        session.FormattedIO.WriteLine("SOUR:FREQ " + val);
                        OldValueFrecuency = val;
                    }


                    // Ожидаем завершения
                    session.FormattedIO.WriteLine("*OPC?");
                    session.FormattedIO.ReadString();

                }
                catch (Exception ex)
                {
                    Reset();
                    throw ex;
                }
                Busy = false;
                return true;
            }

            /// <summary>
            /// Сброс конфигурационных параметров устройства в предопределенное состояние
            /// </summary>
            public void Reset()
            {
                try
                {
                    OldRangeVoltage = "0.1";
                    OldRangeCurrent = "";
                    OldValueFrecuency = "50";
                    IsOutputOn = false;

                    session.Clear();

                    session.FormattedIO.Write("F0V1O0");
                    session.AssertTrigger();

                    Busy = false;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            #endregion
        }
    }
}