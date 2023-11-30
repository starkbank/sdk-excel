using System;
using System.Text;
using EllipticCurve;
using System.Numerics;
using System.Security.Cryptography;


namespace StarkBankExcel.Resources
{
    public class keyGen
    {
        public static string hashPassword(string password, string email)
        {
            string encodedPass = Convert.ToBase64String(Encoding.UTF8.GetBytes(password));

            string encodedEmail = email.ToLower();

            string encodedSalt = "saltDeDev";

            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] beforeHash = sha256.ComputeHash(Encoding.UTF8.GetBytes(encodedPass));
                string hashedPassword = BitConverter.ToString(beforeHash).Replace("-", "").ToLower();

                string sha256Salt = BitConverter.ToString(sha256.ComputeHash(Encoding.UTF8.GetBytes(hashedPassword + ":" + encodedSalt))).Replace("-", "").ToLower();

                string sha256Final = BitConverter.ToString(sha256.ComputeHash(Encoding.UTF8.GetBytes(sha256Salt + ":" + encodedEmail))).Replace("-", "").ToLower();

                return sha256Final;
            }
        }

        public static string cleanEmail(string email)
        {
            string[] values = email.Split('@');

            if (values.Length == 2)
            {
                string name = values[0].Split('+')[0];
                string domain = values[1];
                return (name + '@' + domain).ToLower();
            }
            return "";
        }

        static BigInteger convertToBigInt(string hash)
        {
            return BigInteger.Parse("0" + hash, System.Globalization.NumberStyles.HexNumber);
        }

        static PrivateKey generateNewRandomKey()
        {
            var privateKey = new PrivateKey();
            return privateKey;
        }

        static public PrivateKey generateKeyFromPassword(string password, string email)
        {

            string formattedEmail = cleanEmail(email);

            string hash = hashPassword(password, formattedEmail);

            BigInteger secret = convertToBigInt(hash);

            PrivateKey privateKey = new PrivateKey("secp256k1", secret);

            return privateKey;
        }

        static public string generateSessionAccessId(string session)
        {
            return "session/" + session;
        }

        static public string generateMemberAccessId(string workspaceId, string email)
        {
            return "workspace/" + workspaceId + "/email/" + email;
        }

    }
}
