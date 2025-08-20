// index.js
const express = require('express');
const path = require('path');

const { gerarTermoAnuencia } = require('./termos/termoAnuencia');
const { gerarTermoAditivoRescisao } = require('./termos/termoAditivoRescisao');
const { gerarPortariaPAI } = require('./termos/portariaPAI');   // <-- NOVO

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.get('/health', (_req, res) => res.json({ ok: true }));

app.post('/gerar-termo', async (req, res) => {
  try {
    const { buffer, fileName } = await gerarTermoAnuencia(req.body);
    res.set({
      'Content-Disposition': `attachment; filename="${fileName}"`,
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(err.status || 500).json({ erro: err.message || "Falha ao gerar o Termo de Anuência." });
  }
});

app.post('/gerar-termo-aditivo-rescisao', async (req, res) => {
  try {
    const { buffer, fileName } = await gerarTermoAditivoRescisao(req.body);
    res.set({
      'Content-Disposition': `attachment; filename="${fileName}"`,
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(err.status || 500).json({ erro: err.message || "Falha ao gerar o Termo Aditivo de Rescisão." });
  }
});

// -------- NOVA ROTA: PORTARIA PAI --------
app.post('/gerar-portaria-pai', async (req, res) => {
  try {
    const { buffer, fileName } = await gerarPortariaPAI(req.body);
    res.set({
      'Content-Disposition': `attachment; filename="${fileName}"`,
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(err.status || 500).json({ erro: err.message || "Falha ao gerar a Portaria PAI." });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Servidor rodando em http://localhost:${PORT}`);
});
