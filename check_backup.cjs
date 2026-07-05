const fs = require('fs');
const path = require('path');

try {
  console.log('Listing /app contents:');
  const appContents = fs.readdirSync('/app');
  console.log(appContents);
  
  console.log('Listing /app/applet contents:');
  const appletContents = fs.readdirSync('/app/applet');
  console.log(appletContents);
} catch (err) {
  console.error(err);
}
