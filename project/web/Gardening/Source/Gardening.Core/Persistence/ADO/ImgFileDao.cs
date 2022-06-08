using System;
using System.Collections;
using System.Data;
using Gardening.Core.Domain;
using Gardening.Core.Persistence.ADO.RowMapper;
using Spring.Data.Common;
using Spring.Data.Core;

namespace Gardening.Core.Persistence.ADO
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

        public IList GetByParent(string parentId)
        {
            if (parentId.Trim() == "")
            {
                return null;
            }

            string cmd = "SELECT * FROM IMG_FILE WHERE PARENT_ID = @parentId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("parentId", DbType.String).Value = parentId;

            return AdoTemplate.QueryWithRowMapper(CommandType.Text, cmd, new ImgFileRowMapper(), dbParameters);
        }

        public ImgFile Create(ImgFile file)
        {
            if (file == null)
            {
                throw new ArgumentNullException();
            }

            string cmd = "INSERT INTO IMG_FILE (FILE_ID, PARENT_ID, NAME, URI) VALUES (@FileId, @ParentId, @Name, @Uri)";

            IDbParameters dbParameters = CreateDbParameters();
            file.FileId = Utility.GetGuid();
            dbParameters.Add("FileId", DbType.String).Value = file.FileId;
            dbParameters.Add("ParentId", DbType.String).Value = file.ParentId;
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

        public void DeleteByParent(string parentId)
        {
            string cmd = "DELETE FROM IMG_FILE WHERE PARENT_ID = @parentId";

            IDbParameters dbParameters = CreateDbParameters();
            dbParameters.Add("parentId", DbType.String).Value = parentId;

            AdoTemplate.ExecuteNonQuery(CommandType.Text, cmd, dbParameters);
        }

        #endregion
    }
}
