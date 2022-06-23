using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;
using System.Data;
using System.IO;

namespace WF_Work2
{
    public static class FunctExcel
    {
        private static DataTableCollection tableCollection;
        #region 1. Открытие файла
        public static void OpenFile(string path)
        {
            string result = string.Empty;

            path = path.Replace("\r\n", "");

            var psi = new ProcessStartInfo
            {
                FileName = "explorer",
                Arguments = $"{path}"
            };

            Process.Start(psi);
        }

        /// <summary>
        /// Диалоговое окно на открытие файла
        /// </summary>
        /// <param name="path"></param>
        public static void DialogOpenFile(string path)
        {
            DialogResult result = MessageBox.Show("Файл сохранен!\nОткрыть?", "Сохранение", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (result == DialogResult.Yes)
            {
                Process.Start("explorer", path);
            }
        }

        #endregion
        #region 2. Получить адрес папки/файла/места куда сохранить

        /// <summary>
        /// Выбор папки
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static string ChooseFolder(string path)
        {
            bool isCorrectPath = false;
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();

            try
            {
                folderBrowserDialog1.SelectedPath = path;
                isCorrectPath = folderBrowserDialog1.ShowDialog() == DialogResult.OK;
            }
            catch (Exception)
            {
                folderBrowserDialog1.SelectedPath = "C://";
                isCorrectPath = folderBrowserDialog1.ShowDialog() == DialogResult.OK;
            }

            if (isCorrectPath)
            {
                path = folderBrowserDialog1.SelectedPath;
            }
            return path;
        }

        /// <summary>
        /// Выбрать файл который нужно открыть
        /// </summary>
        /// <param name="filter">фильтр</param>
        /// <returns></returns>
        public static string ChooseFile(string filter = "All files (*.*)|*.*")
        {
            string result = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = filter;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    result = openFileDialog.FileName;
                }
            }
            return result;
        }

        /// <summary>
        /// Выбирает место, куда сохранить файл
        /// </summary>
        /// <param name="filter"></param>
        /// <returns></returns>
        public static string SaveFile(string filter = "All files (*.*)|*.*")
        {
            string result = String.Empty;
            try
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();

                saveFileDialog1.Filter = filter;
                saveFileDialog1.FilterIndex = 2;
                saveFileDialog1.RestoreDirectory = true;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    result = saveFileDialog1.FileName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка сохраниения:\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return result;

        }
        #endregion

        #region 3. Работа с файлами

        /// <summary>
        /// Создать/Перезаписать файл
        /// </summary>
        public static bool CreateAndWrite(string path, string text)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }

                FileStream file = new FileStream(path, FileMode.OpenOrCreate);
                StreamWriter stream = new StreamWriter(file);

                stream.Write(text);

                stream.Close();
                file.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            return true;
        }

        /// <summary>
        /// 3.2 Открыть папку в котором лежит файл
        /// </summary>
        /// <param name="path"></param>
        public static void OpenFolderByFile(string path)
        {
            try
            {
                OpenFile(new DirectoryInfo(path).Parent.FullName);

            }
            catch (Exception)
            {

                throw;
            }
        }

        #endregion
        

    }   
}