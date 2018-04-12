using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrintQuote
{
    class Section
    {
        //holds the section title
        public string title { get; set; }

        //holds the totals for the section
        public double totalMaterialCost { get; private set; }
        public double totalLabourCost { get; private set; }
        public double totalLabourHours { get; private set; }
        public double totalCost { get; private set; }

        //holds the values for each material in the section
        public List<string> _materialTypes = new List<string>();
        public List<string> materialTypes
        {
            get { return _materialTypes; }
            set { _materialTypes = value; }
        }

        public List<int> _quantity = new List<int>();
        public List<int> quantity
        {
            get { return _quantity; }
            set { _quantity = value; }
        }
        public List<double> _materialUnitCosts = new List<double>();
        public List<double> materialUnitCosts
        {
            get { return _materialUnitCosts; }
            set { _materialUnitCosts = value; }
        }
        public List<double> _materialCosts = new List<double>();
        public List<double> materialCosts
        {
            get { return _materialCosts; }
            set { _materialCosts = value; }
        }
        public List<double> _labourUnitCosts = new List<double>();
        public List<double> labourUnitCosts
        {
            get { return _labourUnitCosts; }
            set { _labourUnitCosts = value; }
        }
        public List<double> _labourCosts = new List<double>();
        public List<double> labourCosts
        {
            get { return _labourCosts; }
            set { _labourCosts = value; }
        }

        //used to calculate all the section totals
        public void calc_sectionTotals()
        {
            totalLabourCost = 0;
            totalMaterialCost = 0;
            totalLabourHours = 0;

            foreach (double cost in labourCosts)
            {
                this.totalLabourCost += cost;
            }

            foreach (double cost in materialCosts)
            {
                this.totalMaterialCost += cost;
            }

            totalCost = totalLabourCost + totalMaterialCost;
        }
    }
}
