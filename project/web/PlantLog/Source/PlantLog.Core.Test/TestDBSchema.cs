using System;
using System.Collections;
using NHibernate.Cfg;
using NHibernate.Tool.hbm2ddl;
using NUnit.Framework;
using Spring.Context;
using Spring.Data.Common;
using Spring.Objects.Factory.Config;

namespace PlantLog.Core.Test
{
    [TestFixture]
    public class TestDBSchema
    {
        private string dialectName = string.Empty;
        public static IApplicationContext ctx = Utility.ApplicationContext;

        [Test]
        public void Test_000_ExportSchema()
        {
            string connectionString = GetConnectionString();
            string assemblyName = GetAssemblyName();
            Hashtable t = new Hashtable();
            IDictionary dic = getSpringObjectPropertyValue("SessionFactory", "HibernateProperties") as IDictionary;
            foreach (DictionaryEntry de in dic)
            {
                t.Add(de.Key.ToString(), de.Value.ToString());
                if (de.Key.ToString().Equals("hibernate.dialect"))
                {
                    dialectName = de.Value.ToString();
                }
            }
            //這個不可以改
            t.Add("hibernate.connection.connection_string", connectionString);

            Configuration config = new Configuration();
            config.SetProperties(t);
            config.AddAssembly(assemblyName);

            SchemaExport exporter = new SchemaExport(config);
            exporter.SetOutputFile(buildDDLOutputfileName(dialectName));
            exporter.Drop(false, true);
            exporter.Create(false, true);
        }

        private string GetConnectionString()
        {
            IDbProvider dbp = ctx["DbProvider"] as IDbProvider;
            return dbp.ConnectionString;
        }

        private string GetAssemblyName()
        { return "PlantLog.Core"; }

        private Object getSpringObjectPropertyValue(string objectName, string propertyName)
        {
            IObjectDefinition def =
                ((IConfigurableApplicationContext)ctx).ObjectFactory.GetObjectDefinition(objectName);
            return def.PropertyValues.GetPropertyValue(propertyName).Value;
        }

        private string buildDDLOutputfileName(string dn)
        {
            int a = dn.LastIndexOf('.');
            a = a + 1;
            int b = dn.Length - a;
            string n = dn.Substring(a, b) + ".txt";
            return n;
        }
    }
}
