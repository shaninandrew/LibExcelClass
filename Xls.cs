using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using System.Collections.Generic;
using System.Text;
using ExcelDataReader;
using System.IO;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using System.Text.Json;
                 
/// <summary>
/// Класс для создания объектов производного класса сохраненных в XLS файле.
/// Аналог XML файла/запроса для обмены данными, но с гибкими настройками.
///  Необходим для импорта данных и обработки данных как объекты производьных классов.
///  Поля объектов должын иметь public type var {set;get} !!!
///  Иначе не работает!
/// </summary>
namespace XlsNETFramework
{
   public  class Xls2Class<T>
    {

        public T Result;

        /// <summary>
        ///  Конструктур класса
        /// </summary>
        /// <param name="name">Имя файла для обработки Xls -> список объектов формата колонок. </param>
        /// <param name="Header_Synonyms">Необязательно. Список строк формата X->Y (разделитель ->), 
        /// которые помогают заменить названия  столбцов 
        /// в исходном файле с X на Y для стандартного объекта на Y. 
        /// Условно колонки ФИО->Name.
        /// Также замена работает по номеру колонки: @2->Название  - все 2 колонки будут  изменены на "Название".
        /// Нумерация идет с 1-й колонки (условно A).
        /// </param>
        public Xls2Class(string name, List<string> Header_Synonyms = null)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            FileStream fileStream = new FileStream(name, FileMode.Open, FileAccess.Read);
            ExcelReaderConfiguration configuration = new ExcelReaderConfiguration();
            configuration.AutodetectSeparators = [';'];



            string JSON = "[";

            using (var exldr = ExcelReaderFactory.CreateReader(fileStream, configuration))
            {

                //формат даты для JSON документов
                var format_date_4_json = "yyyy-MM-ddTHH:mm:ssK";

                int index = 0;

                //Заголовок
                List<string> Header = new List<string>();

                //Построчный список
                List<string> Data = new List<string>();

                // Готовим 
                // Json вида: массив из объектов A:dfgkd B:2 C:ferok

                while (exldr.Read())
                {
                    List<string> list = new List<string>();
                    list.Clear();

                    for (int i = 0; i < exldr.FieldCount; i++)
                    {
                        string s = exldr.GetValue(i).ToString();
                        list.Add(s);
                    }


                    //Запоминаем заголовок
                    if (index == 0)
                    {
                        Header = list;

                        //если переименование колонок нужкно - делаем!
                        if (Header_Synonyms != null)
                        {
                            foreach (string syn in Header_Synonyms)
                            {
                                string command = syn.TrimStart();      // удаляем пробелы

                                //комментарии пропускаем
                                if (command.StartsWith("#")) continue;
                                if (command.StartsWith("//")) continue;

                                string[] syn_x = command.Split("->", StringSplitOptions.RemoveEmptyEntries);

                                //строго 2 параметра, что менять на что
                                if (syn_x.Length != 2) continue;

                                //Обходим список заголовков и меняем их
                                for (int i = 0; i < Header.Count; i++)
                                {

                                    //Просто по номеру
                                    if (syn_x[0].Trim() == "@" + (i + 1).ToString())
                                    {
                                        Header[i] = syn_x[1];
                                    }
                                    else //+

                                    //Точное попадание!
                                    // " НаЗваНие" == Название т.к. название==название
                                    if (Header[i].ToLower().Trim() == syn_x[0].ToLower().Trim())
                                    {
                                        Header[i] = syn_x[1];
                                    }
                                    else
                                    //Есть частинчный фрагмент: может быть такая штука aaa и aa, тут конфиг должен быть непротиворечив
                                    if (Header[i].ToLower().IndexOf(syn_x[0].ToLower()) > -1)
                                    {
                                        Header[i] = syn_x[1];
                                    }

                                } //For

                            }   //ForEach

                        }  //обработка синонимов

                    }
                    else
                    {

                        //последнее зпт для продоложение
                        if (index > 1)
                        {
                            JSON += ", ";
                        }

                        JSON += "{ ";

                        Data = list;
                        for (int i = 0; i < Data.Count; i++)
                        {
                            //добавляем с каждой строки данные в свой список по колонек i
                            JSON += "\"" + Header[i] + "\":";


                            //заголовок маленькими буквами
                            string h = Header[i].ToLower();

                            //угадываем тип колонки
                            string guess_type = "text";
                            if (
                                (h.IndexOf("момент") > 0)
                                || (h.IndexOf("срок") > 0)
                                || (h.IndexOf("дата") > 0)
                                || (h.IndexOf("date") > 0)
                                )
                            { guess_type = "date"; }




                            string v = Data[i];
                            string kav = "\"";

                            int check_int = 0;
                            float check_float = 0.0f;
                            double check_double = 0.0;
                            DateTime check_date = DateTime.Now;

                            if (int.TryParse(v, out check_int)) { kav = ""; }
                            else
                            if (float.TryParse(v, out check_float)) { kav = ""; }
                            else
                            if (double.TryParse(v, out check_double)) { kav = ""; }
                            else
                            if ((DateTime.TryParse(v, out check_date)) || (guess_type == "date"))
                            {
                                kav = @"""";
                                v = check_date.ToUniversalTime().ToString(format_date_4_json);
                            }

                            JSON += kav + v + kav;

                            if (i < Data.Count - 1) { JSON += ","; }

                        }
                        //закрытие объектных
                        JSON += " } ";
                    }
                    //индекс
                    index++;


                } //while
                exldr.Close();

            }   //Using

            JSON += " ]";

            // Магия!
            try
            {
                using (JsonDocument jdoc = JsonDocument.Parse(JSON))
                {

                    Result = jdoc.Deserialize<T>();
                }
            }
            catch (Exception ex)
            {
                Debug.Write(ex.Message);
                Debug.Write(ex.StackTrace);
            }

        }

    }
}
