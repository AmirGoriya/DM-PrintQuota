using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrintQuote
{
    class Quote
    {
        //holds all the sections in the quote
        private List<Section> _sections = new List<Section>();
        public List<Section> sections
        {
            get { return _sections; }
            set { _sections = value; }
        }

        //holds the title of the quote
        public string title { get; }

        //holds the totals for the quote
        public double totalMaterialCost { get; private set; }
        public double totalLabourCost { get; private set; }
        public double totalLabourHours { get; private set; }
        public double totalCost { get; private set; }

        //holds the extra costs for the quote
        public double extraCosts { get; set; }

        //holds the cost dedeuctions for the quote
        public double costDedeductions { get; set; }
        public Quote(string title)
        {
            this.title = title;
        }

        public void calcTotals()
        {
            //resets the totals
            totalLabourHours = 0;
            totalLabourCost = 0;
            totalMaterialCost = 0;
            totalCost = 0;

            foreach (Section sec in sections)
            {
                totalLabourHours += sec.totalLabourHours;
                totalMaterialCost += sec.totalMaterialCost;
                totalLabourCost += sec.totalLabourCost;
                totalCost += sec.totalCost;
            }

        }
    }
}
