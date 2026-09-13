#!/usr/bin/env node
// USERS_JSON için parola özeti üretir.
// Kullanım:  node scripts/parola-hash.js <parola>
// Çıkan değeri kullanıcıya "passwordHash" alanı olarak ekleyin,
// "password" alanını silin. Örnek:
//   {"username":"berke","passwordHash":"scrypt:...","name":"Berke","role":"staff"}

const { scryptSync, randomBytes } = require("crypto");

const parola = process.argv[2];
if (!parola) {
  console.error("Kullanım: node scripts/parola-hash.js <parola>");
  process.exit(1);
}

const salt = randomBytes(16).toString("hex");
const hash = scryptSync(parola, salt, 64).toString("hex");
console.log(`scrypt:${salt}:${hash}`);
