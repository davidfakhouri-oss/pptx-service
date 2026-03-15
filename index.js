const express = require('express'); // v3
const app = express();
app.use(express.json({ limit: '10mb' }));

app.get('/', (req, res) => res.send('OK'));

app.post('/generate', async (req, res) => {
  try {
    let code = req.body.code;

    if (typeof code === 'string') {
      // Clean any markdown fences if Claude added them
      code = code.replace(/^```javascript\n?/, '').replace(/^```js\n?/, '').replace(/^```\n?/, '').replace(/```$/, '').trim();
    }

    // Replace the module.exports ending with something we can call
    code = code.replace(
      /module\.exports\s*=\s*async\s*function\s*\(\)\s*\{[\s\S]*?return\s+await\s+pptx\.write\("nodebuffer"\);\s*\}/,
      ''
    );

    // Add PptxGenJS require if not present
    if (!code.includes("require('pptxgenjs')") && !code.includes('require("pptxgenjs")')) {
      code = "const PptxGenJS = require('pptxgenjs');\n" + code;
    }

    // Append the write call
    code += `\nmodule.exports = async function() { return await pptx.write("nodebuffer"); }`;

    // Write code to temp file and execute
    const fs = require('fs');
    const path = require('path');
    const tmpFile = path.join('/tmp', `slide_${Date.now()}.js`);
    fs.writeFileSync(tmpFile, code);

    // Clear require cache and load the module
    delete require.cache[require.resolve(tmpFile)];
    const slideModule = require(tmpFile);
    const buffer = await slideModule();

    // Clean up
    fs.unlinkSync(tmpFile);

    res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.set('Content-Disposition', 'attachment; filename="presentation.pptx"');
    res.send(buffer);

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => console.log(`PPTX service running on port ${PORT}`));
