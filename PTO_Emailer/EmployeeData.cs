namespace PTO_Emailer
{
    class EmployeeData
    {
        private string EmployeeName;
        private string EmployeeFirstName;
        private string EmployeeLastName;
        private string VacationTime;
        private string SickTime;

        public EmployeeData()
        {
            EmployeeName = "";
            VacationTime = "";
            SickTime = "";
        }


        public override string ToString()
        {
            return "Employee Name:\t" + this.EmployeeName + "\r\n" +
                   "Vacation Time:\t" + this.VacationTime + "\r\n" +
                   "Sick Time:    \t" + this.SickTime;
        }


        public string FullName
        {
            get
            {
                return EmployeeName;
            }
            set
            {
                string[] empName = value.Split(',');
                try
                {
                    if (empName[1].Contains("quot;"))
                    {
                        empName[1] = empName[1].Replace("quot;", "\"");
                    }
                    EmployeeFirstName = empName[1];
                    EmployeeLastName = empName[0];
                    EmployeeName = EmployeeLastName + ", " + EmployeeFirstName;
                }
                catch
                {
                    EmployeeFirstName = "Problem";
                    EmployeeLastName = "Parsing";
                    EmployeeName = "Problem Parsing";
                }
            }
        }


        public string FirstName
        {
            get
            {
                return EmployeeFirstName;
            }
        }


        public string LastName
        {
            get
            {
                return EmployeeLastName;
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
