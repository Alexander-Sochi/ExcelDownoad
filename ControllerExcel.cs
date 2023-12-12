 /* foreach (var item in staffingDataFilter)
         {
             *//*  Отрисовка строки "НОМЕР ПО ПОРЯДКУ"*//*
             using (var range = cell["A" + iterationStep + ":A" + (iterationStep + 3)])
             {
                 ApplyBorderAndMerge(range);
             }
             *//*Отрисовка строки "Фамилия,инициалы,должнеость и т.д"*//*
             using (var range = cell["B" + iterationStep + ":E" + (iterationStep + 3)])
             {
                 ApplyBorderAndMerge(range);
             }

             *//*Отрисовка строки "Табкльный номер"*//*
             using (var range = cell["F" + iterationStep + ":F" + (iterationStep + 3)])
             {
                 ApplyBorderAndMerge(range);
             }
             *//* Цикл для отрисовки клеток строки отметка о явках и неявках по числам месяца верхний ряд*//*
             for (int i = 0; i < 16; i++)
             {
                 string range = Convert.ToString(alphabet[alphabetNum]) + iterationStep + ':' + Convert.ToString(alphabet[alphabetNum]) + (iterationStep + 1);
                 using (var rang = cell[range])
                 {
                     ApplyBorderAndMerge(rang);
                     alphabetNum++;
                 }
             }
             //Обнуление alphabetNum до исходного числа 
             alphabetNum = 6;
             *//* Цикл для отрисовки клеток строки отметка о явках и неявках по числам месяца нижний ряд*//*
             for (int i = 0; i < 17; i++)
             {
                 string range = Convert.ToString(alphabet[alphabetNum]) + (iterationStep + 2) + ':' + Convert.ToString(alphabet[alphabetNum]) + (iterationStep + 3);
                 using (var rang = cell[range])
                 {
                     ApplyBorderAndMerge(rang);
                     alphabetNum++;
                 }
             }


             using (var range = cell["W" + iterationStep + ":X" + iterationStep])
             {
                 ApplyBorderAndMerge(range);
             }
             using (var range = cell["Y" + iterationStep + ":Z" + iterationStep])
             {
                 ApplyBorderAndMerge(range);
             }
             using (var range = cell["W" + (iterationStep + 1) + ":X" + (iterationStep + 1)])
             {
                 ApplyBorderAndMerge(range);
             }
             using (var range = cell["Y" + (iterationStep + 1) + ":Z" + (iterationStep + 1)])
             {
                 ApplyBorderAndMerge(range);
             }
             using (var range = cell["W" + (iterationStep + 2) + ":X" + (iterationStep + 3)])
             {
                 ApplyBorderAndMerge(range);
             }

             using (var range = cell["Y" + (iterationStep + 2) + ":Z" + (iterationStep + 3)])
             {
                 ApplyBorderAndMerge(range);
             }
             using (var range = cell[Convert.ToString(alphabet[22]) + iterationStep + ':' + Convert.ToString(alphabet[23]) + iterationStep])
             {
                 ApplyBorderAndMerge(range);
             }
             using (var range = cell["AA" + iterationStep + ":AA" + (iterationStep + 3)])
             {
                 ApplyBorderAndMerge(range);
             }
             using (var range = cell["AB" + iterationStep + ":AC" + (iterationStep + 3)])
             {
                 ApplyBorderAndMerge(range);
             }
             using (var range = cell["AD" + iterationStep + ":AD" + (iterationStep + 3)])
             {
                 ApplyBorderAndMerge(range);
             }
             using (var range = cell["AE" + iterationStep + ":AE" + (iterationStep + 3)])
             {
                 ApplyBorderAndMerge(range);
             }
             using (var range = cell["AF" + iterationStep + ":AG" + (iterationStep + 3)])
             {
                 ApplyBorderAndMerge(range);
             }
             using (var range = cell["AH" + iterationStep + ":AH" + (iterationStep + 3)])
             {
                 ApplyBorderAndMerge(range);
             }
             using (var range = cell["AI" + iterationStep + ":AJ" + (iterationStep + 3)])
             {
                 ApplyBorderAndMerge(range);
             }
             using (var range = cell["AK" + iterationStep + ":AK" + (iterationStep + 3)])
             {
                 ApplyBorderAndMerge(range);
             }
             using (var range = cell["AL" + iterationStep + ":AM" + (iterationStep + 3)])
             {
                 ApplyBorderAndMerge(range);
             }
             using (var range = cell["AN" + iterationStep + ":AO" + (iterationStep + 3)])
             {
                 ApplyBorderAndMerge(range);
             }

             //Отрисовка футтера 


             using (var range = worksheet.Cells["A55:C55"])
             {
                 range.Value = "Ответственное лицо";
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 14; // Размер шрифта 14
                 range.Merge = true;

             }
             //Поля для подписи
             cell["D55:F55"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;


             cell["H55:J55"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

             cell["L55:N55"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

             cell["T55:V55"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

             cell["X55:Z55"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

             cell["AB55:AD55"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

             cell["AF55"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

             cell["AH55:AJ55"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

             cell["AL55:AN55"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

             cell["T59:V59"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

             cell["X59:Z59"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

             cell["AB59:AD59"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

             cell["AF59"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

             cell["AH59:AJ59"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

             cell["AL59:AN59"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

             using (var range = cell["D56:F56"])
             {
                 range.Value = "(должность)";//Текст для столбца
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 9; // Размер шрифта 14
                 range.Merge = true;
                 range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                 range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

             }

             using (var range = cell["H56:J56"])
             {
                 range.Value = "(личная подпись)";//Текст для столбца
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 9; // Размер шрифта 14
                 range.Merge = true;
                 range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                 range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

             }

             using (var range = cell["AB56:AD56"])
             {
                 range.Value = "(расшифровка подписи)";//Текст для столбца
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 9; // Размер шрифта 14
                 range.Merge = true;
                 range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                 range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;


             }

             using (var range = cell["T56:V56"])
             {
                 range.Value = "(должность)";//Текст для столбца
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 9; // Размер шрифта 14
                 range.Merge = true;
                 range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                 range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

             }
             using (var range = cell["X56:Z56"])
             {
                 range.Value = "(личная подпись)";//Текст для столбца
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 9; // Размер шрифта 14
                 range.Merge = true;
                 range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                 range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
             }
             using (var range = cell["AB56:AD56"])
             {
                 range.Value = "(расшифровка подписи)";//Текст для столбца
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 9; // Размер шрифта 14
                 range.Merge = true;
                 range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                 range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
             }


             using (var range = cell["T60:V60"])
             {
                 range.Value = "(должность)";//Текст для столбца
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 9; // Размер шрифта 14
                 range.Merge = true;
                 range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                 range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

             }
             using (var range = cell["X60:Z60"])
             {
                 range.Value = "(личная подпись)";//Текст для столбца
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 9; // Размер шрифта 14
                 range.Merge = true;
                 range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//делает тест по центру по горизонтали 
                 range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
             }
             using (var range = cell["AB60:AD60"])
             {
                 range.Value = "(расшифровка подписи)";//Текст для столбца
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 9; // Размер шрифта 14
                 range.Merge = true;
                 range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//делает тест по центру по горизонтали 
                 range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
             }

             using (var range = cell["AO55"])
             {
                 range.Value = "г";//Текст для столбца
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 9; // Размер шрифта 14
             }
             using (var range = cell["AO59"])
             {
                 range.Value = "г";//Текст для столбца
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 9; // Размер шрифта 14
             }

             using (var range = cell["Q55:R55"])
             {
                 range.Merge = true;
                 range.Value = "Руководитель";//Текст для столбца
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 14; // Размер шрифта 14
             }

             using (var range = cell["Q59"])
             { range.Value = "Работник";//Текст для столбца
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 14; // Размер шрифта 14
             }

             cell["Q56"].Value = "структорного подразделения";


             using (var range = cell["Q56"])
             { range.Value = "Работник";//Текст для столбца
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 10; // Размер шрифта 14
             }
             using (var range = cell["Q60"])
             {
                 range.Value = "кадровой службы";//Текст для столбца
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 10; // Размер шрифта 14
             }
             cell["S59:AN59"].Merge = true;
             cell["A52:AO54"].Merge = true;
      *//*       cell["A56:D56"].Merge = true;*//*
             cell["D55:P55"].Merge = true;
             cell["A57:AO58"].Merge = true;
             cell["A59:P60"].Merge = true;
             cell["AE56:AN56"].Merge = true;
             cell["S55:AN55"].Merge = true;
             cell["A56:C56"].Merge = true;
             cell["AO59"].Merge = true;
             cell["AO56"].Merge = true;
             cell["AO55"].Merge = true;


             //Заполнение данными 


             using (var range = cell["B"+ iterationStep])
             {
                 range.Value = item.FIO;//Текст для столбца
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 9; // Размер шрифта 14  
                 range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//делает тест по центру по горизонтали 
                 range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
             }

             using (var range = cell["A" + iterationStep])
             {
                 range.Value = indexNum;//Текст для столбца
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 9; // Размер шрифта 14  
                 range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//делает тест по центру по горизонтали 
                 range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
             }




             alphabetNum = 6;
             DateTime dt = date;
             double sum = 0;
             int testNumber = 6;
             int numberTest = 6;
             while (dt < dateend)
             {
                 var _day = item.Days.Find(d => d.Date == dt);
                 if (dt.Day <= 15)
                 {
                     using (var range = cell[Convert.ToString(alphabet[testNumber]) + (iterationStep)])
                     {
                         range.Value = _day.Value;//Текст для столбца
                         range.Style.Font.Bold = true; // Жирный шрифт
                         range.Style.Font.Size = 9; // Размер шрифта 14  
                         range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//делает тест по центру по горизонтали 
                         range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                     }
                     testNumber++;
                 }


                 if (dt.Day > 15)
                 {
                     using (var range = cell[Convert.ToString(alphabet[numberTest]) + (iterationStep + 2)])
                     {
                         range.Value = _day.Value;//Текст для столбца
                         range.Style.Font.Bold = true; // Жирный шрифт
                         range.Style.Font.Size = 9; // Размер шрифта 14  
                         range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//делает тест по центру по горизонтали 
                         range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                     }
                     numberTest++;
                 }



                 sum += _day.Value ?? 0;


                 dt = dt.AddDays(1);
             }

             using (var range = cell["Y" + (iterationStep + 2)])
             {
                 range.Value = sum;//Текст для столбца
                 range.Style.Font.Bold = true; // Жирный шрифт
                 range.Style.Font.Size = 9; // Размер шрифта 14  
                 range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//делает тест по центру по горизонтали 
                 range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
             }


             iterationStep += 4;
             alphabetNum = 6;
             indexNum++;


             if (dt.Day == 15)
             {
                 testNumber = 6;
             }
             if (dt.Day == 16)
             {
                 numberTest= 6;
             }

         }*/
