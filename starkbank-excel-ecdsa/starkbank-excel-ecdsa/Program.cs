using System.Net;
using System.Numerics;


namespace EllipticCurve
{

    internal class Program
    {

        static void Main(string[] args)
        {

            if (args.Length < 1) { return; }
            string functionSelection = args[0];

            if (functionSelection == "generatePrivateKeySecret")
            {
                PrivateKey privateKey = new PrivateKey();
                Console.WriteLine(privateKey.secret);
            }
            if (args.Length == 1) { return; }

            if (functionSelection == "getPublicKeyFromSecret")
            {
                BigInteger secret = BigInteger.Parse(args[1]);

                PrivateKey privateKeyObject = new PrivateKey("secp256k1", secret);
                PublicKey publicKey = privateKeyObject.publicKey();
                Console.WriteLine(publicKey.toPem());
            }

            if (functionSelection == "getSecretFromString")
            {
                string secretString = args[1];

                string secretHash = Ecdsa.sha256(secretString);
                BigInteger secret = Utils.BinaryAscii.numberFromHex(secretHash);
                PrivateKey privateKey = new PrivateKey("secp256k1", secret);
                Console.WriteLine(privateKey.secret);
            }
            if (args.Length == 2) { return; }

            if (functionSelection == "sign")
            {
                string message = args[1];
                string secret = args[2];

                BigInteger secretInteger = BigInteger.Parse(secret);
                PrivateKey privateKeyObject = new PrivateKey("secp256k1", secretInteger);
                Signature signature = Ecdsa.sign(message, privateKeyObject);
                Console.WriteLine(signature.toBase64());
            }

            if (functionSelection == "signPrivateKey")
            {
                string message = args[1];
                string privateKey = args[2];

                PrivateKey privateKeyObject = PrivateKey.fromPem(privateKey);
                Signature signature = Ecdsa.sign(message, privateKeyObject);
                Console.WriteLine(signature.toBase64());
            }
        }
    }
}
