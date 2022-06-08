using System.Data;
using Spring.Data;
using PlantLog.Core.Domain;

namespace PlantLog.Core.Persistence.ADO.RowMapper
{
    class OwnerRowMapper : IRowMapper
    {
        public object MapRow(IDataReader dataReader, int rowNum)
        {
            Owner owner = new Owner();
            owner.OwnerId = dataReader.GetString(0);
            owner.DisplayName = dataReader.GetString(1);
            if (dataReader.IsDBNull(2))
            {
                owner.Nickname = null;
            }
            else
            {
                owner.Nickname = dataReader.GetString(2);
            }
            owner.Topic = dataReader.GetString(3);
            if (dataReader.IsDBNull(4))
            {
                owner.Email = null;
            }
            else
            {
                owner.Email = dataReader.GetString(4);
            }
            if (dataReader.IsDBNull(5))
            {
                owner.Description = null;
            }
            else
            {
                owner.Description = dataReader.GetString(5);
            }
            owner.IsApprove = dataReader.GetBoolean(6);
            owner.CreatorId = dataReader.GetString(7);
            owner.CreateDateTime = dataReader.GetDateTime(8);
            owner.ModifierId = dataReader.GetString(9);
            owner.ModifyDateTime = dataReader.GetDateTime(10);

            return owner;
        }
    }
}
