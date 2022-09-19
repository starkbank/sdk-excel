﻿

namespace EllipticCurve.Utils {

    public static class Base64 {

        public static byte[] decode(string base64String) {
            return Convert.FromBase64String(base64String);
        }

        public static string encode(byte[] bytes) {
            return Convert.ToBase64String(bytes);
        }
    }
}
