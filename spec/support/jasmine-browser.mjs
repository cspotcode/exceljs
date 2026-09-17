export default {
  srcDir: 'dist',
  srcFiles: ['exceljs.js'],
  specDir: 'spec/browser',
  specFiles: ['exceljs.spec.js'],
  helpers: [],
  env: {
    stopSpecOnExpectationFailure: false,
    stopOnSpecFailure: false,
    random: false,
  },
  listenAddress: 'localhost',
  hostname: 'localhost',
  browser: {
    name: 'headlessChrome',
  },
};
