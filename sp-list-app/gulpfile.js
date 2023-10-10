const gulp = require('gulp');

gulp.task('set-sp-site', async function() {
  var cfgObj = 'export const cfg = { SP_SITE: "';
  var site = "";
  var iSite = process.argv.indexOf("--site");
  if(iSite > -1) {
      site = process.argv[iSite + 1];
  }
  cfgObj = cfgObj + site + '"';
  var sslFlag = true;
  var iSsl = process.argv.indexOf("--ssl");
  if(iSsl >-1) {
      sslFlag = process.argv[iSsl + 1];
  }
  cfgObj = cfgObj + ", SSL: " + sslFlag;
  var listId = "";
  var iList = process.argv.indexOf("--listid");
  if(iList >-1) {
    listId = process.argv[iList + 1];
  }
  cfgObj = cfgObj + ', LIST_ID: "' + listId + '" }';

  return require('fs').writeFileSync('./src/app-config.ts', cfgObj);
});


const build = require('@microsoft/sp-build-web');

build.configureWebpack.mergeConfig({
  build:{
    //add anything here if needed
  }
});
build.initialize(gulp);
