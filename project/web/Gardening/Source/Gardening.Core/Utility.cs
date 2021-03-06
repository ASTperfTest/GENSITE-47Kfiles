using System;
using Spring.Context;
using Spring.Context.Support;
using Gardening.Core.Service;

namespace Gardening.Core
{
    public class Utility
    {
        private static IApplicationContext ctx = null;

        public static IApplicationContext ApplicationContext
        {
            get
            {
                if (ctx == null)
                {
                    ctx = ContextRegistry.GetContext();
                }
                return ctx;
            }
            set
            {
                ctx = value;
            }
        }

        public static string GetGuid()
        {
            return Guid.NewGuid().ToString().ToLower();
        }

        public static string GetUploadPath()
        {
            return "";
        }
    }
}
