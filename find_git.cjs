const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

console.log('Current working directory:', process.cwd());

let dir = process.cwd();
while (dir !== path.dirname(dir)) {
  if (fs.existsSync(path.join(dir, '.git'))) {
    console.log('Found .git in:', dir);
    try {
      const status = execSync('git status', { cwd: dir, encoding: 'utf-8' });
      console.log('Git Status:\n', status);
    } catch (e) {
      console.log('Git Status failed:', e.message);
    }
    break;
  }
  dir = path.dirname(dir);
}
