using System;
using System.Collections.Generic;
using System.Data.OracleClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutomateMappingTool
{
    class Validation
    {
        DataGridView dataGridView;
        OracleConnection ConnectionProd;
        OracleConnection ConnectionTemp;
        OracleCommand cmd;
        OracleDataReader reader;

        string outputPath;
        string type;
        ListBox listBox = new ListBox();
        List<int> lstIndex;
        string message = null, suffixMkt = null, code = null, month = null, channel = null, mkt = null, order = null,
            province = null, effective = null, expire = null, entry = null, install = null,
            downSpeed = null, upSpeed = null, uom = null;

        private List<string> lstSpeed = new List<string>();
        private List<string> lstFormatSpeed = new List<string>();
        private List<string> lstProv = new List<string>();
        private List<string> lstEff = new List<string>();
        private List<string> lstExp = new List<string>();
        private List<string> lstMonth = new List<string>();
        private List<string> lstUOM = new List<string>();
        private List<string> lstEntry = new List<string>();
        private List<string> lstInstall = new List<string>();
        private List<string> lstProdType = new List<string>();
        private List<string> lstChannel = new List<string>();
        private bool isCorrect;

        public Validation(DataGridView dgv, OracleConnection connProd, OracleConnection connTemp, string output)
        {
            this.dataGridView = dgv;
            this.ConnectionProd = connProd;
            this.ConnectionTemp = connTemp;
            this.outputPath = output;
        }

        public List<int> indexDgv
        {
            get
            {
                return this.lstIndex;
            }
            set
            {
                this.lstIndex = value;
            }
        }

        /// <summary>
        /// Process for validate data of requirement file
        /// </summary>
        /// <returns></returns>
        public ListBox verify(string t)
        {
            isCorrect = true;
            lstSpeed = new List<string>();
            lstProv = new List<string>();
            lstEff = new List<string>();
            lstExp = new List<string>();
            lstMonth = new List<string>();
            listBox.Items.Clear();
            lstIndex = new List<int>();
            type = t;

            clearSelection();

            for (int i = 0; i < dataGridView.RowCount; i++)
            {
                if (type.Equals("VAS"))
                {
                    code = dataGridView.Rows[i].Cells[0].Value.ToString().Trim();
                    channel = dataGridView.Rows[i].Cells[1].Value.ToString();
                    mkt = dataGridView.Rows[i].Cells[2].Value.ToString().ToUpper();
                    order = dataGridView.Rows[i].Cells[3].Value.ToString();
                    string speed = dataGridView.Rows[i].Cells[5].Value.ToString().ToUpper().Trim();
                    uom = getUOM(speed);
                    province = dataGridView.Rows[i].Cells[6].Value.ToString().ToUpper();
                    effective = dataGridView.Rows[i].Cells[7].Value.ToString().Trim();
                    expire = dataGridView.Rows[i].Cells[8].Value.ToString().Trim();

                    verifySpeed(i, speed);
                    verifyProv(i);
                    verifyDate(i);
                }
                else if (type.Equals("Disc"))
                {
                    code = dataGridView.Rows[i].Cells[0].Value.ToString().Trim();
                    month = dataGridView.Rows[i].Cells[1].Value.ToString();
                    channel = dataGridView.Rows[i].Cells[2].Value.ToString();
                    mkt = dataGridView.Rows[i].Cells[3].Value.ToString().ToUpper();
                    order = dataGridView.Rows[i].Cells[4].Value.ToString();
                    string speed = dataGridView.Rows[i].Cells[6].Value.ToString().ToUpper().Trim();
                    uom = getUOM(speed);
                    province = dataGridView.Rows[i].Cells[7].Value.ToString().ToUpper();
                    effective = dataGridView.Rows[i].Cells[8].Value.ToString();
                    expire = dataGridView.Rows[i].Cells[9].Value.ToString();

                    verifyGroupMonth(i);
                    verifySpeed(i, speed);
                    verifyProv(i);
                    verifyDate(i);
                }
                else
                {
                    mkt = dataGridView.Rows[i].Cells[0].Value.ToString().ToUpper();
                    string speed = dataGridView.Rows[i].Cells[1].Value.ToString().ToUpper().Trim();
                    if (speed.Contains('/'))
                    {
                        string[] spSpeed = speed.Split('/');
                        downSpeed = spSpeed[0].Trim();
                        upSpeed = spSpeed[1].Trim();

                        uom = getUOM(downSpeed);
                        if (String.IsNullOrEmpty(uom))
                        {
                            uom = getUOM(upSpeed);

                            if (String.IsNullOrEmpty(uom))
                            {
                                //error not found uom
                            }
                        }
                    }
                    else
                    {
                        string msg = "";

                        lstSpeed.Add(msg);
                        listBox.Items.Add(msg);
                        lstIndex.Add(i);

                        hilightRow(type, "speed", i);

                        isCorrect = false;
                    }

                    order = dataGridView.Rows[i].Cells[4].Value.ToString();
                    channel = dataGridView.Rows[i].Cells[5].Value.ToString();
                    effective = dataGridView.Rows[i].Cells[8].Value.ToString();
                    expire = dataGridView.Rows[i].Cells[9].Value.ToString();
                    entry = dataGridView.Rows[i].Cells[10].Value.ToString().Trim();
                    install = dataGridView.Rows[i].Cells[11].Value.ToString().Trim();

                    verifySpeed(i, downSpeed);
                }

            }

            if (lstSpeed.Count > 0)
            {
                message += "[" + DateTime.Now.ToString() + "] >>Mismatch speed between suffix MKT Code and download speed<<" + "\r\n";

                foreach (string msg in lstSpeed)
                {
                    message += "    " + msg + "\r\n";
                }
            }

            if (lstProv.Count > 0)
            {
                message += "[" + DateTime.Now.ToString() + "] " + "\r\n";

                foreach (string msg in lstProv)
                {
                    message += "    " + msg + "\r\n";
                }
            }

            if (lstEff.Count > 0)
            {
                message += "[" + DateTime.Now.ToString() + "] >>Wrong effective date format or effective is null<<" + "\r\n";

                foreach (string msg in lstEff)
                {
                    message += "    " + msg + "\r\n";
                }
            }

            if (lstExp.Count > 0)
            {
                message += "[" + DateTime.Now.ToString() + "] >>Wrong expire date format<<" + "\r\n";

                foreach (string msg in lstExp)
                {
                    message += "    " + msg + "\r\n";
                }
            }

            if (lstMonth.Count > 0)
            {
                message += "[" + DateTime.Now.ToString() + "] >>Invalid GroupID<<" + "\r\n";

                foreach (string msg in lstMonth)
                {
                    message += "    " + msg + "\r\n";
                }
            }

            if (message != "")
            {
                message += "\r\n" + "***********************************************" + "\r\n";

                string strFilePath = outputPath + "\\Validation_Log" + ".txt";
                using (StreamWriter writer = new StreamWriter(strFilePath, true))
                {
                    writer.Write(message);
                }
            }
            // return isCorrect;
            return listBox;
        }

        private void verifySpeed(int index, string speed)
        {
            string suffixMkt = "";

            if (mkt.Contains("-"))
            {
                string[] lstmkt = mkt.Split('-');
                suffixMkt = lstmkt[1].Trim();
                int n, speedID;
                if (int.TryParse(suffixMkt, out n))
                {
                    speedID = Convert.ToInt16(suffixMkt);

                    if (uom == "G")
                    {
                        speedID = Convert.ToInt32(suffixMkt) * 1000;
                    }

                    string query = "SELECT SPEED_ID, SPEED_DESC FROM HISPEED_SPEED WHERE SPEED_ID = " + speedID;

                    cmd = new OracleCommand(query, ConnectionProd);
                    reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        reader.Read();
                        int speed_desc = Convert.ToInt32(reader["SPEED_DESC"]);

                        if (uom == "G")
                        {
                            speed_desc = speed_desc / 1024000;
                        }
                        else if (uom == "M")
                        {
                            speed_desc = speed_desc / 1024;
                        }

                        suffixMkt = Convert.ToString(speed_desc);
                    }
                    else
                    {
                        string msg = code + ", " + month + " month, MKTCode : " + mkt + ", Speed " + speed +
                            " >> Not found the speed in database[HiSPEED_SPEED]";

                        lstSpeed.Add(msg);
                        listBox.Items.Add(msg);
                        lstIndex.Add(index);

                        hilightRow(type, "speed", index);

                        isCorrect = false;
                    }

                    reader.Close();
                }
                else
                {
                    string msg = code + ", " + month + " month, MKTCode : " + mkt + ", Speed " + speed +
                            " >> Wrong format MKT Code";
                    lstSpeed.Add(msg);
                    listBox.Items.Add(msg);
                    lstIndex.Add(index);

                    hilightRow(type, "speed", index);

                    isCorrect = false;
                }
            }
            else if (mkt.Equals("ALL"))
            {
                suffixMkt = "ALL";
            }

            if (speed.Equals("ALL") == false)
            {
                speed = Regex.Replace(speed, "[^0-9]", "");
            }

            if (speed.Equals(suffixMkt) == false)
            {
                string msg = code + ", " + month + " month, MKTCode : " + mkt;
                lstSpeed.Add(msg);
                listBox.Items.Add(msg);
                lstIndex.Add(index);

                hilightRow(type, "speed", index);

                isCorrect = false;
            }
        }

        private bool verifyUOM(int index)
        {
            int num;
            bool hasUOM;


            if ((int.TryParse(downSpeed, out num)) && (int.TryParse(upSpeed, out num)))
            {
                string msg = mkt + ", Speed : " + "Don't have UOM of Speed";
                lstUOM.Add(msg);
                listBox.Items.Add(msg);
                lstIndex.Add(index);

                hilightRow(type, "speed", index);

                isCorrect = false;
                hasUOM = false;
            }
            else
            {
                hasUOM = true;
            }

            return hasUOM;
        }

        private string getUOM(string speed)
        {
            int num;
            string uom = null;

            if (int.TryParse(speed, out num) == false)
            {
                uom = Regex.Replace(speed, "[0-9]", "");
            }

            return uom;
        }

        private void verifyProv(int index)
        {
            //Check Province
            if (province.Contains(","))
            {
                if (province.Contains("ALL"))
                {
                    string msg = code + ", " + month + " month, MKTCode : " + mkt + ", Order : " + order +
                        ", province : " + province + " >>> Conflict province";
                    hilightRow(type, "province", index);
                    lstProv.Add(msg);
                    listBox.Items.Add(msg);
                    lstIndex.Add(index);
                    isCorrect = false;
                }
                else
                {
                    string[] lstProvince = province.Split(',');

                    foreach (string prov in lstProvince)
                    {
                        bool hasRow = GetProvince(prov);
                        if (hasRow == false)
                        {
                            string msg = code + ", " + month + " month, MKTCode : " + mkt + ", Order : " + order + ", province : " + province;
                            lstProv.Add(msg);
                            listBox.Items.Add(msg);
                            lstIndex.Add(index);

                            hilightRow(type, "province", index);

                            isCorrect = false;
                        }
                    }
                }
            }
            else
            {
                bool hasRow;
                if (province.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    province = "ALL";
                    hasRow = true;
                }
                else
                {
                    hasRow = GetProvince(province);
                }

                if (hasRow == false)
                {
                    string msg = code + ", " + month + " month, MKTCode : " + mkt + ", Order : " + order + ", province : " + province;
                    lstProv.Add(msg);
                    listBox.Items.Add(msg);
                    lstIndex.Add(index);

                    hilightRow(type, "province", index);

                    isCorrect = false;
                }
            }

        }

        //Check province in DB
        private bool GetProvince(string province)
        {
            string queryProv = "SELECT * FROM DISCOUNT_CRITERIA_PROVINCE WHERE DP_TYPE = '" + province.Trim() + "'";
            bool hasRow = true;
            try
            {
                cmd = new OracleCommand(queryProv, ConnectionProd);
                reader = cmd.ExecuteReader();
                reader.Read();
                if (reader.HasRows == false)
                {
                    hasRow = false;
                }

                reader.Close();
            }
            catch (Exception ex)
            {
                string log = "[" + DateTime.Now.ToString() + "]" + "Cannot searching province from database(DISCOUNT_CRITERIA_PROVINCE)." +
                    "\r\n" + "Error message : " + ex.Message;

                hasRow = false;
            }

            return hasRow;
        }

        private void verifyDate(int index)
        {
            ChangeFormat chgFormat = new ChangeFormat();
            //Check Date
            string dateEff = chgFormat.formatDate(effective);

            if (dateEff == "Invalid")
            {
                hilightRow(type, "effective", index);

                string msg = code + ", " + month + " month, MKTCode : " + mkt;
                lstEff.Add(msg);

                listBox.Items.Add(msg);
                lstIndex.Add(index);

                isCorrect = false;
            }

            if (expire.Equals("-"))
            {
                expire = String.Empty;
            }

            if (String.IsNullOrEmpty(expire) == false)
            {
                string date = chgFormat.formatDate(expire);
                if (date == "Invalid")
                {
                    hilightRow(type, "expire", index);

                    string msg = code + ", " + month + " month, MKTCode : " + mkt;
                    lstExp.Add(msg);
                    listBox.Items.Add(msg);
                    lstIndex.Add(index);

                    isCorrect = false;
                }
            }
        }

        private void verifyGroupMonth(int index)
        {
            List<Dictionary<string, string>> lstRangeM = FormatGroupMonth(month);
            Dictionary<string, string> dicresultM = new Dictionary<string, string>();
            Dictionary<string, string> dicGroupID = new Dictionary<string, string>();

            for (int num = 0; num < lstRangeM.Count; num++)
            {
                dicresultM = lstRangeM[num];
                string minMonth = dicresultM["min" + num];
                string maxMonth = dicresultM["max" + num];

                //Check GroupID
                string queryGroup = "SELECT * FROM discount_criteria_group where DG_DISCOUNT = '" + code + "' and dg_month_min = " +
                    minMonth + " and dg_month_max = " + maxMonth;

                try
                {
                    cmd = new OracleCommand(queryGroup, ConnectionProd);
                    reader = cmd.ExecuteReader();
                    reader.Read();
                    if (reader.HasRows)
                    {
                        string distinct = code + minMonth + maxMonth;
                        string dgGroup = reader["DG_GROUPID"].ToString();

                        if (dicGroupID.ContainsKey(distinct) == false)
                        {
                            dicGroupID.Add(distinct, dgGroup);
                        }
                    }
                    else
                    {
                        hilightRow(type, "month", index);

                        string msg = code + ", month : " + minMonth + "-" + maxMonth + ", MKTCode : " + mkt;
                        lstMonth.Add(msg);
                        listBox.Items.Add(msg);
                        lstIndex.Add(index);

                        isCorrect = false;
                    }

                    reader.Close();
                }
                catch (Exception ex)
                {
                    string log = "[" + DateTime.Now.ToString() + "]" + "Cannot searching discount group(DG_GROUPID) from database(DISCOUNT_CRITERIA_GROUP)." +
                        "\r\n" + "Error message : " + ex.Message;
                }

            }
        }

        public List<Dictionary<string, string>> FormatGroupMonth(string month)
        {
            List<Dictionary<string, string>> lstRangeM = new List<Dictionary<string, string>>();
            Dictionary<string, string> dic;

            if (month.Equals("ตลอดอายุการใช้งาน"))
            {
                month = "-1";
            }

            if (month.Contains(","))
            {
                string[] lstMonth = month.Split(',');

                for (int j = 0; j < lstMonth.Length; j++)
                {
                    dic = new Dictionary<string, string>();

                    if (lstMonth[j].StartsWith("-1"))
                    {
                        if (lstMonth[j].Length > 2)
                        {
                            lstMonth[j] = lstMonth[j].Substring(2);

                            string[] lstRange = lstMonth[j].Split('-');
                            dic.Add("min" + j, "-1");
                            dic.Add("max" + j, lstRange[1]);
                            lstRangeM.Add(dic);
                        }
                        else
                        {
                            dic.Add("min" + j, lstMonth[j]);
                            dic.Add("max" + j, lstMonth[j]);
                            lstRangeM.Add(dic);
                        }
                    }
                    else
                    {
                        if (lstMonth[j].Contains("-"))
                        {
                            string[] lstRange = lstMonth[j].Split('-');
                            dic.Add("min" + j, lstRange[0]);
                            dic.Add("max" + j, lstRange[1]);
                            lstRangeM.Add(dic);
                        }
                        else
                        {
                            dic.Add("min" + j, lstMonth[j]);
                            dic.Add("max" + j, lstMonth[j]);
                            lstRangeM.Add(dic);
                        }
                    }
                }
            }
            else
            {
                dic = new Dictionary<string, string>();

                if (month.StartsWith("-1"))
                {
                    if (month.Length > 2)
                    {
                        month = month.Substring(2);
                        string[] lstRange = month.Split('-');

                        dic.Add("min0", "-1");
                        dic.Add("max0", lstRange[1]);
                    }
                    else
                    {
                        dic.Add("min0", month);
                        dic.Add("max0", month);
                    }
                }
                else
                {
                    if (month.Contains("-"))
                    {
                        string[] lstRange = month.Split('-');
                        dic.Add("min0", lstRange[0]);
                        dic.Add("max0", lstRange[1]);
                    }
                    else
                    {
                        dic.Add("min0", month);
                        dic.Add("max0", month);
                    }
                }

                lstRangeM.Add(dic);
            }

            return lstRangeM;
        }

        private void verifyContract(int index)
        {
            string queryEnt = "SELECT * FROM TRUE9_BPT_CONTRACT WHERE ENTRY = '" + entry + "'";
            string queryIns = "SELECT * FROM TRUE9_BPT_CONTRACT WHERE INSTALL = '" + install + "'";

            //Entry Code
            cmd = new OracleCommand(queryEnt, ConnectionTemp);
            reader = cmd.ExecuteReader();
            if (reader.HasRows == false)
            {
                string msg = "Entry Code " + entry + " of " + mkt + " ,Not found in table >> TRUE9_BPT_CONTRACT";
                lstEntry.Add(msg);
                listBox.Items.Add(msg);
                lstIndex.Add(index);

                hilightRow(type, "entry", index);

                reader.Close();
            }

            //Install code
            cmd = new OracleCommand(queryIns, ConnectionTemp);
            reader = cmd.ExecuteReader();
            if (reader.HasRows == false)
            {
                string msg = "Install Code " + install + " of " + mkt + " ,Not found in table >> TRUE9_BPT_CONTRACT";

                lstInstall.Add(msg);
                listBox.Items.Add(msg);
                lstIndex.Add(index);

                hilightRow(type, "install", index);

                reader.Close();
            }
        }

        private void verifyProdType(int index)
        {
            string prefixMKT;
            if (mkt.StartsWith("TRL"))
            {
                prefixMKT = mkt.Substring(0, 5);
            }
            else
            {
                prefixMKT = mkt.Substring(0, 2);
            }

            string query = "SELECT * FROM TRUE9_BPT_HISPEED_PRODTYPE WHERE MKT = '" + prefixMKT + "'";
            cmd = new OracleCommand(query, ConnectionTemp);
            reader = cmd.ExecuteReader();

            if (reader.HasRows == false)
            {
                string msg = "The prefix " + prefixMKT + " of " + mkt + " ,Not found in table >> TRUE9_BPT_HISPEED_PRODTYPE";
                lstProdType.Add(msg);
                listBox.Items.Add(msg);
                lstIndex.Add(index);

                hilightRow(type, "mkt", index);

                reader.Close();
            }
        }

        /// <summary>
        /// Check channel from file compare with channel in database
        /// </summary>
        /// <param name="channel"></param>
        /// <returns></returns>
        private List<string> verifyChannel(string channel)
        {
            List<string> lstInvaid = new List<string>();

            if (channel.Contains(","))
            {
                string[] lstCh = channel.Split(',');

                foreach (string val in lstCh)
                {
                    val.Trim();

                    if (val.Equals("ALL", StringComparison.OrdinalIgnoreCase) == false)
                    {
                        if (lstChannel.Contains(val) == false)
                        {
                            lstInvaid.Add(val);
                        }
                    }

                }
            }
            else
            {
                if (channel.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    channel = "DEFAULT";
                }

                if (lstChannel.Contains(channel) == false)
                {
                    lstInvaid.Add(channel);
                }
            }

            return lstInvaid;
        }

        /// <summary>
        /// Get channel from DB
        /// </summary>
        /// <returns>List of channel in DB</returns>
        public List<string> GetChannelFromDB()
        {
            List<string> lstChannelFromDB = new List<string>();

            try
            {
                //Get all channel in DB
                string query = "SELECT DISTINCT(SALE_CHANNEL) FROM HISPEED_CHANNEL_PROMOTION";
                cmd = new OracleCommand(query, ConnectionProd);
                reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    lstChannelFromDB.Add(reader["SALE_CHANNEL"].ToString());
                }

                reader.Close();
            }
            catch (Exception ex)
            {
                string msg = "Cannot get sale channel from database. Please click validate file again or check internet connection.";

                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                lstChannelFromDB = new List<string>();
            }

            return lstChannelFromDB;
        }

        /// <summary>
        /// Clear selected row 
        /// </summary>
        private void clearSelection()
        {
            dataGridView.ClearSelection();

            for (int i = 0; i < dataGridView.RowCount; i++)
            {
                for (int j = 0; j < dataGridView.ColumnCount; j++)
                {
                    dataGridView.Rows[i].Cells[j].Style.BackColor = Color.Empty;
                }
            }
        }

        /// <summary>
        /// Hilight mistake row
        /// </summary>
        /// <param name="indexRow"></param>
        /// <param name="indexCol"></param>
        private void hilightRow(string type, string key, int indexRow)
        {
            Dictionary<string, int> indexDisc = new Dictionary<string, int>
            { {"month",1}, {"channel",2 },{"mkt",3},{"order",4},{"speed",6},{"province",7},{"effective",8},{"expire",9} };

            Dictionary<string, int> indexVas = new Dictionary<string, int>
            {{"channel",1 },{"mkt",2},{"order",3},{"speed",5},{"province",6},{"effective",7},{"expire",8} };

            Dictionary<string, int> indexHis = new Dictionary<string, int>
            {{"mkt",0 },{"speed",1},{"order",4},{"channel",5},{"effective",8},{"expire",9},{"entry",10}, {"install",11} };

            if (type.Equals("VAS"))
            {
                int indexCol = indexVas[key];
                dataGridView.Rows[indexRow].Cells[indexCol].Style.BackColor = Color.Red;
            }
            else if (type.Equals("Disc"))
            {
                int indexCol = indexDisc[key];
                dataGridView.Rows[indexRow].Cells[indexCol].Style.BackColor = Color.Red;
            }
            else
            {
                int indexCol = indexDisc[key];
                dataGridView.Rows[indexRow].Cells[indexCol].Style.BackColor = Color.Red;
            }

        }
    }
}
