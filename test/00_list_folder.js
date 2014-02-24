var 
  ExchangeCalendarList = require('../calendar')
  , config = require('../config.json')
;

var calendar = new ExchangeCalendarList();

calendar
.request({
  url: config.url,
  username: config.username,
  password: config.password,
  folderid: config.folderid,
  maximum: 1000
})
.then(function(items) {
  console.log(items);
}, function(err) {
  console.log(err);
});
