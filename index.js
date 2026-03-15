const express = require('express'); // v5
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

    const fs = require('fs');
    const path = require('path');
    const tmpFile = path.join('/tmp', `slide_${Date.now()}.js`);

    fs.writeFileSync(tmpFile, code);
    delete require.cache[require.resolve(tmpFile)];
    const slideModule = require(tmpFile);
    const buffer = await slideModule();
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
