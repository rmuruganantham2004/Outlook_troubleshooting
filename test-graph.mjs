import fetch from "node-fetch";

const token = process.argv[2];

const res = await fetch("https://graph.microsoft.com/v1.0/me", {
  headers: {
    Authorization: `Bearer ${token}`,
  },
});

console.log(res.status);
console.log(await res.text());