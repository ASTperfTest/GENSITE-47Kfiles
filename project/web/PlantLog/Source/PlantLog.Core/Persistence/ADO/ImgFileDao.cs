using System;
using System.Collections;
using System.Data;
using PlantLog.Core.Domain;
using PlantLog.Core.Persistence.ADO.RowMapper;
using Spring.Data.Common;
using Spring.Data.Core;

namespace PlantLog.Core.Persistence.ADO
{
    public class ImgFileDao : AdoDaoSupport, IImgFileDao
    {
        #region ISourceFileDao Members

        public ImgFile Get(string fileId)
        {
            if (fileId.Trim() == "")
            {
                return null;
            }

            string cmd = "SELECT * FROM IMG_FILE WHERE FILE_ID = @fileId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("fileId", DbType.String).Value = fileId;

            IList result = AdoTemplate.QueryWithRowMapper(CommandType.Text, cmd, new ImgFileRowMapper(), dbParameters);

            if (result.Count > 0)
            {
                return (ImgFile)result[0];
            }
            else
            {
                return null;
            }
        }

        public IList GetByEntry(string entryId)
        {
            if (entryId.Trim() == "")
            {
                return null;
            }

            string cmd = "SELECT * FROM IMG_FILE WHERE ENTRY_ID = @entryId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("entryId", DbType.String).Value = entryId;

            return AdoTemplate.QueryWithRowMapper(CommandType.Text, cmd, new ImgFileRowMapper(), dbParameters);
        }

        public ImgFile Create(ImgFile file)
        {
            if (file == null)
            {
                throw new ArgumentNullException();
            }

            string cmd = "INSERT INTO IMG_FILE (FILE_ID, ENTRY_ID, NAME, URI) VALUES (@FileId, @EntryId, @Name, @Uri)";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("FileId", DbType.String).Value = file.FileId;
            dbParameters.Add("EntryId", DbType.String).Value = file.EntryId;
            dbParameters.Add("Name", DbType.String).Value = file.Name;
            dbParameters.Add("Uri", DbType.String).Value = file.Uri;

            AdoTemplate.ExecuteNonQuery(CommandType.Text, cmd, dbParameters);
            return Get(file.FileId);
        }

        public ImgFile Update(ImgFile file)
        {
            if (file == null)
            {
                throw new ArgumentNullException();
            }

            string cmd = "UPDATE IMG_FILE SET NAME = @Name, URI = @Uri WHERE FILE_ID = @FileId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("FileId", DbType.String).Value = file.FileId;
            dbParameters.Add("Name", DbType.String).Value = file.Name;
            dbParameters.Add("Uri", DbType.String).Value = file.Uri;

            AdoTemplate.ExecuteNonQuery(CommandType.Text, cmd, dbParameters);
            return Get(file.FileId);
        }

        public void Delete(string fileId)
        {
            string cmd = "DELETE FROM IMG_FILE WHERE FILE_ID = @FileId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("FileId", DbType.String).Value = fileId;

            AdoTemplate.ExecuteNonQuery(CommandType.Text, cmd, dbParameters);
        }

        public void DeleteByEntry(string entryId)
        {
            string cmd = "DELETE FROM IMG_FILE WHERE ENTRY_ID = @entryId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("entryId", DbType.String).Value = entryId;

            AdoTemplate.ExecuteNonQuery(CommandType.Text, cmd, dbParameters);
        }

        #endregion
    }
}
