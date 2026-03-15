const express = require('express'); // v8
const PptxGenJS = require('pptxgenjs');
const app = express();
app.use(express.json({ limit: '10mb' }));

app.get('/', (req, res) => res.send('OK'));

app.post('/generate', async (req, res) => {
  try {
    let code = req.body.code;

    if (typeof code === 'string') {
      code = code
        .replace(/^```javascript\n?/, '')
        .replace(/^```js\n?/, '')
        .replace(/^```\n?/, '')
        .replace(/```$/, '')
        .trim();
    }

    // Strip # from hex colors just in case
    code = code.replace(/'#([0-9A-Fa-f]{6})'/g, "'$1'");
    code = code.replace(/"#([0-9A-Fa-f]{6})"/g, '"$1"');

    // Remove the require('pptxgenjs') line — we inject it directly
    code = code.replace(/const PptxGenJS\s*=\s*require\(['"]pptxgenjs['"]\)\s*;?\n?/g, '');

    // Remove module.exports line — we handle the write ourselves
    code = code.replace(/module\.exports\s*=\s*async\s*function\s*\(\)\s*\{[\s\S]*?\}\s*;?\s*$/m, '');

    // Build the final executable code
    const finalCode = `
const PptxGenJS = require('pptxgenjs');
${code}
module.exports = async function() { return await pptx.write("nodebuffer"); };
`;

    const fs = require('fs');
    const path = require('path');

    // Write to /app/tmp so it can access node_modules
    const tmpDir = path.join('/app', 'tmp');
    if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir);
    const tmpFile = path.join(tmpDir, `slide_${Date.now()}.js`);

    fs.writeFileSync(tmpFile, finalCode);

    let slideModule;
    try {
      delete require.cache[require.resolve(tmpFile)];
      slideModule = require(tmpFile);
    } catch (syntaxErr) {
      try { fs.unlinkSync(tmpFile); } catch(e) {}
      return res.status(500).json({ 
        error: syntaxErr.message,
        type: 'SYNTAX_ERROR',
        line: syntaxErr.stack ? syntaxErr.stack.split('\n')[0] : 'unknown'
      });
    }

    let buffer;
    try {
      buffer = await slideModule();
    } catch (runtimeErr) {
      try { fs.unlinkSync(tmpFile); } catch(e) {}
      return res.status(500).json({ 
        error: runtimeErr.message,
        type: 'RUNTIME_ERROR',
        line: runtimeErr.stack ? runtimeErr.stack.split('\n')[1] : 'unknown'
      });
    }

    try { fs.unlinkSync(tmpFile); } catch(e) {}

    res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.set('Content-Disposition', 'attachment; filename="presentation.pptx"');
    res.send(buffer);

  } catch (err) {
    res.status(500).json({ 
      error: err.message,
      type: 'GENERAL_ERROR'
    });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => console.log(`PPTX service running on port ${PORT}`));
