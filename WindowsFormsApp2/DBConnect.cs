using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

using System.Configuration;
namespace WindowsFormsApp2
{
    class DBConnect
    {
        private  MySqlConnection connection;
        private string server;
        private string database;
        private string uid;
        private string password;
        private string port;
        private string timeout;
    

        //Constructor
        public DBConnect()
        {
            Initialize();
        }

        public DBConnect(string server, string database, string uid, string password,string port,string timeout) {
        
            string connectionString;
            connectionString = "SERVER=" + server + ";" + "Port=" + port + ";" + "DATABASE=" +
            database + ";" + "UID=" + uid + ";" + "PASSWORD=" + password + ";" + "Connection Timeout=" + timeout + ";";

            connection = new MySqlConnection(connectionString);
           
        }

        


        //Initialize values
        private void Initialize()
        {
            string appPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string configFile = System.IO.Path.Combine(appPath, System.Reflection.Assembly.GetExecutingAssembly().GetName().Name + ".exe.config");


            ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap();
            configFileMap.ExeConfigFilename = configFile;
            Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);
         
            server = config.AppSettings.Settings["Server"].Value;
            database = config.AppSettings.Settings["Database"].Value;
            uid = config.AppSettings.Settings["Uid"].Value;
            password = config.AppSettings.Settings["Password"].Value;
            port=config.AppSettings.Settings["Port"].Value;
            timeout= config.AppSettings.Settings["Timeout"].Value;
            string connectionString;
            connectionString = "SERVER=" + server + ";" + "Port=" + port + ";" + "DATABASE=" +
            database + ";" + "UID=" + uid + ";" + "PASSWORD=" + password + ";"+ "Connection Timeout=" + timeout + ";";

            connection = new MySqlConnection(connectionString);
        }
       

        public bool OpenConnection()
        {
            try
            {
                connection.Open();
                return true;
            }
            catch (MySqlException ex)
            {
              
                switch (ex.Number)
                {
                    case 0:
                        MessageBox.Show("Cannot connect to server.  Contact administrator");
                        break;

                    case 1045:
                        MessageBox.Show("Invalid username/password, please try again");
                        break;
                }
                return false;
            }
        }

        //Close connection
        public bool CloseConnection()
        {
            try
            {
                connection.Close();
                return true;
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }
        public string fetchMail(string company_codice_fiscale)
        {
            string query = "select email from beneficiaries where company_codice_fiscale=@Id";
            string email ="";

            //Open Connection
            if (this.OpenConnection() == true)
            {
                //Create Mysql Command
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@Id", company_codice_fiscale);
                cmd.Prepare();
                //ExecuteScalar will return one value
                email = (cmd.ExecuteScalar() + "").ToString();

                //close Connection
                this.CloseConnection();

                return email;
            }
            else
            {
                return email;
            }
        }

        public void updateMailAddress(string company_codice_fiscale,string email)
        {
            string query = "UPDATE beneficiaries SET email=@email WHERE company_codice_fiscale=@Id";

            //Open connection
            if (this.OpenConnection() == true)
            {
                //create mysql command
                MySqlCommand cmd = new MySqlCommand();
                //Assign the query using CommandText
                cmd.Parameters.AddWithValue("@Id", company_codice_fiscale);
                cmd.Parameters.AddWithValue("@email", email);
                cmd.CommandText = query;
                //Assign the connection using Connection
                cmd.Connection = connection;

                //Execute query
                cmd.ExecuteNonQuery();

                //close connection
                this.CloseConnection();

           
            }
        }

    }
}
