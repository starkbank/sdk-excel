﻿using System.Text;
using System.Numerics;
using System.Security.Cryptography;


namespace EllipticCurve {

    public static class Ecdsa {

        public static Signature sign(string message, PrivateKey privateKey) {
            string hashMessage = sha256(message);
            BigInteger numberMessage = Utils.BinaryAscii.numberFromHex(hashMessage);
            CurveFp curve = privateKey.curve;
            BigInteger randNum = Utils.Integer.randomBetween(BigInteger.One, curve.N - 1);
            Point randSignPoint = EcdsaMath.multiply(curve.G, randNum, curve.N, curve.A, curve.P);
            BigInteger r = Utils.Integer.modulo(randSignPoint.x, curve.N);
            BigInteger s = Utils.Integer.modulo((numberMessage + r * privateKey.secret) * (EcdsaMath.inv(randNum, curve.N)), curve.N);

            return new Signature(r, s);
        }

        public static bool verify(string message, Signature signature, PublicKey publicKey) {
            string hashMessage = sha256(message);
            BigInteger numberMessage = Utils.BinaryAscii.numberFromHex(hashMessage);
            CurveFp curve = publicKey.curve;
            BigInteger sigR = signature.r;
            BigInteger sigS = signature.s;

            if (sigR < 1 || sigR >= curve.N) {
                return false;
            }
            if (sigS < 1 || sigS >= curve.N) {
                return false;
            }

            BigInteger inv = EcdsaMath.inv(sigS, curve.N);

            Point u1 = EcdsaMath.multiply(
                curve.G,
                Utils.Integer.modulo((numberMessage * inv), curve.N),
                curve.N,
                curve.A,
                curve.P
            );
            Point u2 = EcdsaMath.multiply(
                publicKey.point,
                Utils.Integer.modulo((sigR * inv), curve.N),
                curve.N,
                curve.A,
                curve.P
            );
            Point v = EcdsaMath.add(
                u1,
                u2,
                curve.A,
                curve.P
            );
            if (v.isAtInfinity()) {
                return false;
            }
            return Utils.Integer.modulo(v.x, curve.N) == sigR;
        }

        public static string sha256(string message) {
            byte[] bytes;

            using (SHA256 sha256Hash = SHA256.Create()) {
                bytes = sha256Hash.ComputeHash(Encoding.UTF8.GetBytes(message));
            }

            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < bytes.Length; i++) {
                builder.Append(bytes[i].ToString("x2"));
            }

            return builder.ToString();
        }
    }
}
