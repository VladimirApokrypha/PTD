using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace FuncionalPTD.FunctionalClasses
{
    public class Serializer
    {
        public void SerializeObject(string fileName, object objToSerialize)
        {
            try
            {
                FileStream fstream = File.Open(fileName, FileMode.Create);
                BinaryFormatter binaryFormatter = new BinaryFormatter();
                binaryFormatter.Serialize(fstream, objToSerialize);
                fstream.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
        }

        public object DeserializeObject(string fileName)
        {
            object objToSerialize = null;
            try
            {
                FileStream fstream = File.Open(fileName, FileMode.Open);
                BinaryFormatter binaryFormatter = new BinaryFormatter();
                objToSerialize = binaryFormatter.Deserialize(fstream);
                fstream.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
            return objToSerialize;
        }
    }
}
