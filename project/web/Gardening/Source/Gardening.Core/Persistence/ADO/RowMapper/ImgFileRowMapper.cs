using System.Data;
using Gardening.Core.Domain;
using Spring.Data;

namespace Gardening.Core.Persistence.ADO.RowMapper
{
    class ImgFileRowMapper : IRowMapper
    {
        public object MapRow(IDataReader dataReader, int rowNum)
        {
            ImgFile file = new ImgFile();
            file.FileId = dataReader.GetString(0);
            file.ParentId = dataReader.GetString(1);
            file.Name = dataReader.GetString(2);
            file.Uri = dataReader.GetString(3);

            return file;
        }
    }
}
