using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Net.NetworkInformation;

namespace WordPartition
{
    class NetUtil
    {
        public static string Connect(string ip)
        {
            string result = string.Empty;
            try
            {
                IPAddress ipAddress = IPAddress.Parse(ip);
                IPAddress[] ipadressArr = Dns.GetHostAddresses(ip);
                if (ipadressArr != null && ipadressArr.Length > 0)
                {
                    ipAddress = ipadressArr[0];

                    Ping pingSender = new Ping();
                    PingOptions options = new PingOptions();

                    options.DontFragment = true;

                    string data = "a";
                    byte[] buffer = Encoding.ASCII.GetBytes(data);
                    int timeout = 120;
                    PingReply reply = pingSender.Send(ipAddress, timeout, buffer, options);

                    if (reply.Status == IPStatus.Success)
                    {

                    }
                    else
                    {
                        result = "서버와의 Ping 응답이 없습니다";
                    }
                }
                else
                {
                    result = "ip 정보가 잘못 되었습니다.";
                }
            }
            catch
            {
                result = "ip 정보가 잘못 되었습니다.";
            }
            return result;
        }
    }
}
