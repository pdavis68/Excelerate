using System;
using System.Collections.Generic;
using System.Threading;

namespace Excelerate
{
    public class Relationships
    {
        private int _id_ctr = 0;

        private List<Relationship> _relationships = new List<Relationship>();
        public Relationships()
        {
            
        }

        public string AddRelationship(string type, string target)
        {
            var rel = new Relationship(Interlocked.Increment(ref _id_ctr), type, target);
            _relationships.Add(rel);
            return rel.Id;
        }

        internal string GenerateRelationshipsXml()        
        {
            string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + Environment.NewLine;
            xml += "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
            foreach(var relationship in _relationships)
            {
                xml += relationship.GenerateRelationshipXml();
            }
            xml += "</Relationships>";
            return xml;
        }

        private class Relationship
        {
            public Relationship(int id, string type, string target)
            {
                _id = "rId" + id.ToString();
                _type = type;
                _target = target;
            }

            private string _id;
            public string Id { get => _id; }
            private string _type;
            public string Type { get => _type; }
            private string _target;
            public string Target { get => _target; }
            
            public string GenerateRelationshipXml()
            {
                return $"<Relationship Id=\"{_id}\" Type=\"{_type}\" Target=\"{_target}\"/>";
            }
        }

    }
}