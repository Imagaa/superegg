// File: api/config.js
export default function handler(req, res) {
  // Mengambil URL dari Environment Variable Vercel
  const gasUrl = process.env.GAS_API_URL;
  
  if (!gasUrl) {
    return res.status(500).json({ error: "Environment variable GAS_API_URL belum disetting di Vercel!" });
  }

  res.status(200).json({ url: gasUrl });
}