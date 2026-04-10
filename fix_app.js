import fs from 'fs';
let content = fs.readFileSync('App.tsx', 'utf-8');
content = content.replace(/const PYTHON_CODE = `[\s\S]*?`;/, "import { PYTHON_CODE } from './pythonTemplate';");
content = content.replace("import { ParsedSheetData } from './types';", "import { ParsedSheetData, ParsedWorkbookData } from './types';");
fs.writeFileSync('App.tsx', content);
