using System;
using System.Collections.Generic;
using System.Text;

namespace SharedProject1
{
    public static class Serializer
    {
        /// <summary>
        /// Serialize an object to JSON
        /// </summary>
        /// <typeparam name="T">Object type to serialize</typeparam>
        /// <param name="thingy">Object to get</param>
        /// <returns>serialized string</returns>
        public static string SerializeToJson<T>(T thingy)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T));
                serializer.WriteObject(memoryStream, thingy);
                return Encoding.UTF8.GetString(memoryStream.ToArray());
            }
        }



        /// <summary>
        /// Deserialize an object from JSON
        /// </summary>
        /// <typeparam name="T">Object type to deserialize to</typeparam>
        /// <param name="json">String to deserialize</param>
        /// <returns>Stream to return</returns>
        public static T DeserializeFromJson<T>(string json)
        {
            try
            {
                if (string.IsNullOrEmpty(json))
                {
                    return default(T);
                }
                using (Stream memoryStream = StreamFromString(json))
                {
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T));
                    T result = (T)serializer.ReadObject(memoryStream);
                    return result;
                }
            }
            catch (Exception ex)
            {

                throw new Exception("Error while Deserialisation." + ex.Message + "||" + ex.InnerException.Message);
            }
        }



        /// <summary>
        /// Get stream from a string, helper function to Serializer class
        /// </summary>
        /// <param name="s">string to get stream from</param>
        /// <returns>Stream to return</returns>
        private static Stream StreamFromString(string s)
        {
            Stream stream = null;
            Stream tempStream = null;
            try
            {
                stream = new MemoryStream();
                StreamWriter writer = new StreamWriter(stream);
                writer.Write(s);
                writer.Flush();
                stream.Position = 0;
                tempStream = stream;
                stream = null;
            }
            finally
            {
                if (stream != null)
                {
                    stream.Close();
                }
            }



            return tempStream;
        }
    }

}
