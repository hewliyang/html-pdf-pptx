# html-pdf-pptx

a html -> pdf -> pptx pipeline

## features

1. vectorizing `<i>` fontawesome icons into <svgs> with proper paths so that it converts to a vector graphic instead of being rasterized
2. respects original font definition (if available on the system) instead of auto inferring Noto Sans
3. better positioning of text/VGs when interpolating different font styles in the middle of paragraphs
4. OK for commercial use, apache 2.0 (Puppeteer) & GPLv3 (Libreoffice)

## tradeoffs
slow, but can be parallelized because the individual html -> pdf conversion can be mapreduced

## examples

### Current Manus Export (PPTX)
<img width="1667" height="1083" alt="Screenshot 2025-08-04 at 2 52 25 PM" src="https://github.com/user-attachments/assets/3dbc7c7d-dc13-49c3-966d-f043dbf6c757" />

### This pipeline
<img width="1667" height="1083" alt="Screenshot 2025-08-04 at 2 52 19 PM" src="https://github.com/user-attachments/assets/dbb6f56f-2201-44f9-8243-f17096cd4d0d" />


## manus

`slides-to-pptx.js` supports passing in the Manus `slides.json`, which is also used when requesting for export on the platform

```yaml
POST https://api.manus.im/session.v1.SessionPublicService/CreateSessionFileConvertTask
convertType: "SESSION_FILE_CONVERT_TYPE_HTML_TO_PPT"
fileName: "AWS Costs Breakdown"
# presigned url w/ to the slides.json
fromUrl: "https://private-us-east-1.manuscdn.com/sessionFile/mIaN0vMStA0DPbw0NkboBy/sandbox/imxsodzji... 
```

usage:

```bash
npm i
node slides-to-pptx.js slides.json
```

## prereqs
- libreoffice
- puppeteer and chromium
