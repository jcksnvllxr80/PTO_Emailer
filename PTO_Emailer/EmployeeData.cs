namespace PTO_Emailer
{
    class EmployeeData
    {
        private string EmployeeName;
        private string VacationTime;
        private string SickTime;

        public EmployeeData()
        {
            this.EmployeeName = "";
            this.VacationTime = "";
            this.SickTime = "";
        }


        public override string ToString()
        {
            return "Employee Name:\t" + this.EmployeeName + "\r\n" +
                   "Vacation Time:\t" + this.VacationTime + "\r\n" +
                   "Sick Time:    \t" + this.SickTime;
        }


        public string Name
        {
            get
            {
                return EmployeeName;
            }
            set
            {
                EmployeeName = value;
            }
        }


        public string Vacation
        {
            get
            {
                return VacationTime;
            }

            set
            {
                VacationTime = value;
            }
        }


        public string Sick
        {
            get
            {
                return SickTime;
            }

            set
            {
                SickTime = value;
            }
        }
    }
}
