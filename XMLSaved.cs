using System;
using System.Xml;
using System.IO;

/// <summary>
/// Summary description for Class1
/// </summary>
namespace System.Xml
{
    [Serializable]
    public class XmlSaved<T>
    {
        /// <summary>
        ///     Сохранение структуры в файл
        /// </summary>
        /// <param name="file">Полный путь к файлу</param>
        /// <param name="obj">Структура</param>
        public static void Save(string file, T obj)
        {
            System.Xml.Serialization.XmlSerializer xs = new System.Xml.Serialization.XmlSerializer(typeof(T));
            StreamWriter writer = File.CreateText(file);
            xs.Serialize(writer, obj);
            writer.Flush();
            writer.Close();
        }

        /// <summary>
        ///     Подключение структуры из файла
        /// </summary>
        /// <param name="file">Полный путь к файлу</param>
        /// <returns>Структура</returns>
        public static T Load(string file)
        {
            System.Xml.Serialization.XmlSerializer xs = new System.Xml.Serialization.XmlSerializer(typeof(T));
            StreamReader reader = File.OpenText(file);
            T c = (T)xs.Deserialize(reader);
            reader.Close();
            return c;
        }

        /// <summary>
        ///     Подключение структуры из файла
        /// </summary>
        /// <param name="file">Полный путь к файлу</param>
        /// <returns>Структура</returns>
        public static T LoadFile(string file) { return Load(file); }

        /// <summary>
        ///     Подключение структуры из запроса URL
        /// </summary>
        /// <param name="url">Ссылка</param>
        /// <returns>Структура</returns>
        public static T LoadURL(string url)
        {
            System.Xml.Serialization.XmlSerializer xs = new System.Xml.Serialization.XmlSerializer(typeof(T));
            System.Net.WebRequest wr = System.Net.HttpWebRequest.Create(url);
            wr.Method = "GET";
            wr.Timeout = 30000;
            System.Net.WebResponse rp = wr.GetResponse();
            System.IO.Stream ss = rp.GetResponseStream();
            T c = (T)xs.Deserialize(ss);
            ss.Close();
            rp.Close();
            return c;
        }

        /// <summary>
        ///     Получение папки, из которой запущено приложение
        /// </summary>
        /// <returns>Полный путь к папки с \ на конце</returns>
        public static string GetCurrentDir()
        {
            string fname = System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase.ToString();
            fname = fname.Replace("file:///", "");
            fname = fname.Replace("/", @"\");
            fname = fname.Substring(0, fname.LastIndexOf(@"\") + 1);
            return fname;
        }

        /// <summary>
        ///     Подключение объекта класса по интерфейсу из DLL
        ///     Конструктор класса должен быть без аргументов
        /// </summary>
        /// <param name="filename">Полный путь к файлу</param>
        /// <returns>Объект класса</returns>
        public static T LoadFromDLL(string filename)
        {
            System.Reflection.Assembly asm = System.Reflection.Assembly.LoadFile(filename);
            Type[] tps = asm.GetTypes();
            Type asmType = null;
            foreach (Type tp in tps) if (tp.GetInterface(typeof(T).ToString()) != null) asmType = tp;

            System.Reflection.ConstructorInfo ci = asmType.GetConstructor(new Type[] { });
            return (T)ci.Invoke(new object[] { });
        }

        /// <summary>
        ///     Подключение объекта класса по интерфейсу из DLL по URL
        ///     Конструктор класса должен быть без аргументов
        ///     (Используется для интерфейсов)
        /// </summary>
        /// <param name="url">Ссылка</param>
        /// <returns>Объект класса</returns>
        public static T LoadFromDLL_URL(string url)
        {
            System.Net.WebRequest wr = System.Net.HttpWebRequest.Create(url);
            wr.Method = "GET";
            wr.Timeout = 30000;
            System.Net.WebResponse rp = wr.GetResponse();
            System.IO.Stream ss = rp.GetResponseStream();

            string dd = System.Environment.SpecialFolder.ApplicationData.ToString() + @"\#tmplda\";
            if (!System.IO.Directory.Exists(dd)) System.IO.Directory.CreateDirectory(dd);
            string ff = dd + System.DateTime.Now.Ticks.ToString() + ".dll";

            System.IO.FileStream fs = new FileStream(ff, FileMode.CreateNew);

            int rb = -1;
            while ((rb = ss.ReadByte()) >= 0) fs.WriteByte((byte)rb);
            ss.Close();
            fs.Close();
            rp.Close();

            System.Reflection.Assembly asm = System.Reflection.Assembly.LoadFile(ff);
            Type[] tps = asm.GetTypes();
            Type asmType = null;
            foreach (Type tp in tps) if (tp.GetInterface(typeof(T).ToString()) != null) asmType = tp;

            System.Reflection.ConstructorInfo ci = asmType.GetConstructor(new Type[] { });
            return (T)ci.Invoke(new object[] { });
        }
    }
}
