# Stark Bank Excel Ecdsa Executable

## Overview

This is an Ecdsa Library using the Dotnet framework compiled as an executable.

## Sample Code

### Generate a random private key in secret format (BigInteger in a string format)

```shell
generatePrivateKeySecret
```

### Generate a private key in secret format (BigInteger in a string format) from a secret string

```shell
getSecretFromString secretString
```

### Generate a public key in PEM format from a private key in secret format (BigInteger in a string format)

```shell
getPublicKeyFromSecret secret
```

### Sign a string message using a private key in secret format (BigInteger in a string format)

```shell
sign message secret
```
