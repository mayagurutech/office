const fs = require('fs');
const path = require('path');

// Ensure dist directory exists
const distDir = path.join(__dirname, 'dist');
if (!fs.existsSync(distDir)) {
  fs.mkdirSync(distDir, { recursive: true });
}

// Read Index.html
const indexPath = path.join(__dirname, 'Index.html');
let indexContent = fs.readFileSync(indexPath, 'utf8');

// Regex to find <?!= include('FILENAME'); ?> or <?!= include("FILENAME"); ?>
const includeRegex = /<\?!=\s*include\(['"]([^'"]+)['"]\);\s*\?>/g;

indexContent = indexContent.replace(includeRegex, (match, filename) => {
  const filePath = path.join(__dirname, `${filename}.html`);
  if (fs.existsSync(filePath)) {
    console.log(`Including: ${filename}.html`);
    return fs.readFileSync(filePath, 'utf8');
  } else {
    console.warn(`Warning: File not found for include: ${filename}.html`);
    return `<!-- Include Error: ${filename}.html not found -->`;
  }
});

// Save to dist/index.html
fs.writeFileSync(path.join(distDir, 'index.html'), indexContent, 'utf8');
console.log('Build complete! Unified index.html written to dist/index.html');
