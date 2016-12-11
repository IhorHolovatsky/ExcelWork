using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace COFCO.Forms
{
    public class Constants
    {
        #region Error Messages

        public const string ErrorMessage = "Помилка";

        public const string ParamsInputErrorMessage =
            "Перевірте правильність вводу. Всі колонки повинні бути заповнені та не має бути продубльованих рядків.";

        public const string InputExcelErrorMessage = "Сталась проблема в створенні проміжного Excel файлу. Перевірте вхідний файл та введені дані.";

        public const string InputFileAndDirectoryExistanceErrorMessage = "Вкажіть шлях до вхідного файлу (Файл постачальника) та проміжної папки.";

        public const string TempFileAndDirectoryExistanceErrorMessage = "Вкажіть шлях до файлу з контрактами та вихідної папки.";

        public const string InputFileExistanceErrorMessage = "Вкажіть шлях до вхідного файлу (Файл постачальника).";
        #endregion
    }
}
