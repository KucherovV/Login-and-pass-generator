using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Practice2018Boikov
{
    class LoginAndPass
    {
        public string Login { get; set; }
        public string Pass { get; set; }
        public string NameSurname { get; set; }

        public LoginAndPass(string login, string pass, string nameSurname)
        {
            Login = login;
            Pass = pass;
            NameSurname = nameSurname;
        }

    }
}
