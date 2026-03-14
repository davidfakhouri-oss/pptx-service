const express = require('express');
const PptxGenJS = require('pptxgenjs');
const app = express();
app.get('/', (req, res) => res.send('OK'));

app.post('/generate', async (req, res) => {
  try {
    let parsed = req.body.parsed;

    if (typeof parsed === 'string') {
      parsed = parsed.replace(/^```json\n?/, '').replace(/^```\n?/, '').replace(/```$/, '').trim();
      const jsonStart = parsed.indexOf('{');
      const jsonEnd = parsed.lastIndexOf('}');
      parsed = JSON.parse(parsed.substring(jsonStart, jsonEnd + 1));
    }

    const COLORS = {
      darkNavy: "0D1B2A", accentGold: "C8952A", lightGray: "F2F4F7",
      mediumGray: "8395A7", white: "FFFFFF", textDark: "1A1A2E",
    };

    const ppt = new PptxGenJS();
    ppt.layout = "LAYOUT_WIDE";

    for (let i = 0; i < parsed.slides.length; i++) {
      const slideData = parsed.slides[i];
      const slide = ppt.addSlide();
      const slideNum = i + 1;
      const totalSlides = parsed.slides.length;

      slide.addShape(ppt.ShapeType.rect, { x:0, y:0, w:"100%", h:"100%", fill:{color:COLORS.white}, line:{color:COLORS.white} });
      slide.addShape(ppt.ShapeType.rect, { x:0, y:0, w:0.18, h:"100%", fill:{color:COLORS.darkNavy}, line:{color:COLORS.darkNavy} });
      slide.addShape(ppt.ShapeType.rect, { x:0.18, y:0, w:"100%", h:1.55, fill:{color:COLORS.darkNavy}, line:{color:COLORS.darkNavy} });
      slide.addShape(ppt.ShapeType.rect, { x:0.18, y:1.55, w:13.17, h:0.045, fill:{color:COLORS.accentGold}, line:{color:COLORS.accentGold} });
      slide.addShape(ppt.ShapeType.rect, { x:0.18, y:6.9, w:13.17, h:0.6, fill:{color:COLORS.lightGray}, line:{color:COLORS.lightGray} });
      slide.addShape(ppt.ShapeType.rect, { x:0.18, y:6.88, w:13.17, h:0.04, fill:{color:COLORS.accentGold}, line:{color:COLORS.accentGold} });

      slide.addText(parsed.title.toUpperCase(), { x:0.35, y:6.92, w:9, h:0.3, fontSize:7, color:COLORS.mediumGray, align:"left", charSpacing:1.5 });
      slide.addText(`${slideNum} / ${totalSlides}`, { x:10.5, y:6.92, w:2.8, h:0.3, fontSize:7, color:COLORS.mediumGray, align:"right" });
      slide.addText(slideData.keyMessage, { x:0.35, y:0.18, w:12.8, h:0.9, fontSize:26, bold:true, color:COLORS.white, fontFace:"Calibri", valign:"middle" });
      slide.addText(slideData.subtitle, { x:0.35, y:1.05, w:12.8, h:0.45, fontSize:13, color:COLORS.accentGold, fontFace:"Calibri", italic:true });

      const content = slideData.content;
      if (content.type === "bullets" || content.type === "steps") {
        const bullets = content.items.map((item, idx) => ({
          text: content.type === "steps" ? `${idx+1}.  ${item}` : `•  ${item}`,
          options: { fontSize:15, color:COLORS.textDark, breakLine:true, paraSpaceAfter:8 }
        }));
        slide.addText(bullets, { x:0.5, y:1.75, w:12.85, h:5, fontFace:"Calibri", valign:"top" });
      } else if (content.type === "two-column") {
        const leftItems = content.left.map(item => ({ text:`•  ${item}`, options:{fontSize:15, color:COLORS.textDark, breakLine:true, paraSpaceAfter:8} }));
        slide.addShape(ppt.ShapeType.rect, { x:0.5, y:1.75, w:6.2, h:4.9, fill:{color:COLORS.lightGray}, line:{color:COLORS.lightGray} });
        slide.addText(leftItems, { x:0.65, y:1.9, w:5.9, h:4.6, fontFace:"Calibri", valign:"top" });
        const rightItems = content.right.map(item => ({ text:`•  ${item}`, options:{fontSize:15, color:COLORS.textDark, breakLine:true, paraSpaceAfter:8} }));
        slide.addShape(ppt.ShapeType.rect, { x:6.95, y:1.75, w:6.2, h:4.9, fill:{color:COLORS.lightGray}, line:{color:COLORS.lightGray} });
        slide.addText(rightItems, { x:7.1, y:1.9, w:5.9, h:4.6, fontFace:"Calibri", valign:"top" });
      }
    }

    const buffer = await ppt.write("nodebuffer");
    res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.set('Content-Disposition', 'attachment; filename="presentation.pptx"');
    res.send(buffer);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => console.log(`PPTX service running on port ${PORT}`));
