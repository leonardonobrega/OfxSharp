using System;
using System.Collections.Generic;
using System.Xml;

namespace OfxSharp
{
    public class OfxDocument
    {
        public DateTime StatementStart { get; set; }

        public DateTime StatementEnd { get; set; }

        public AccountType AccType { get; set; }

        public string Currency { get; set; }

        public SignOn SignOn { get; set; }

        public Account Account { get; set; }

        public Balance Balance { get; set; }

        public List<Transaction> Transactions { get; set; }
        public string OriginalHeader { get; set; }
        public XmlDocument Xml { get; set; }

        public override string ToString()
        {
            return $"{OriginalHeader}\n\n{Xml?.OuterXml}";
        }
    }
}