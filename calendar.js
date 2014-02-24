var
  fs = require('fs')
  , promise = require('node-promise')
	, EventEmitter = require('events').EventEmitter
	, util = require('util')
  , https = require('https')
  , xml2js = require('xml2js')
;

var ExchangeCalendarList = function(opts) {
	var self = this;
	EventEmitter.call(this);

  this.options = opts;

  this.ready = false;
  this.templates = {};
  this.loadTemplates(function() {
    self.ready = true;
    self.emit('ready');
  });
};
util.inherits(ExchangeCalendarList, EventEmitter);

exports = module.exports = function(opts) {
	if (!opts) opts = {};
	return new ExchangeCalendarList(opts);
};

ExchangeCalendarList.prototype.loadTemplates = function(callback) {
  var self = this;
  promise.all(fs.readdirSync(__dirname + '/ews_templates').map(function(fn) {
    if (!fn.match(/\.xml$/)) return;

    var p = promise.Promise();

    fs.readFile(__dirname + '/ews_templates/' + fn, 'utf8', function(err, data) {
      if (err)
        p.reject(err);
      else {
        self.templates[fn.substring(0, fn.length-4)] = data;
        p.resolve();
      }
    });

    return p;
  }))
  .then(callback)
  ;
};

ExchangeCalendarList.prototype.request = function(opts) {
  var p = promise.Promise();

  var self = this;
  if (!this.ready) {
    this.once('ready', function() {
      self
        .request(opts)
        .then(function(result) {
          p.resolve(result);
        }, function(err) {
          p.reject(err);
        })
      ;
    });
    return p;
  }

  var url = require('url').parse(opts.url);
  url.method = 'POST';
  url.auth = [opts.username, opts.password].join(':');
  url.headers = {
    'Content-Type': 'text/xml'
  };

  opts.fields = opts.fields ? opts.fields : [
      "item:Subject",
      "item:Body",
      "calendar:Location",
      "calendar:Start",
      "calendar:End",
      "item:Categories"
    ];

  var fields = '';
  opts.fields.forEach(function(field) {
    fields += '<t:FieldURI FieldURI="' + field + '"/>';
  });

  var xml;

  if (opts.start) {
    var start = new Date(opts.start);
    var end = new Date(opts.end);
  
    start.toString = end.toString = function() {
      return this.toISOString().replace(/\..+/, '') + '-00:00';
    };

    xml = self.templates.ews_find_items
      .replace('$FolderId$', opts.folderid)
      .replace('$StartDate$', start)
      .replace('$EndDate$', end)
    ;
  }
  else {
    xml = self.templates.ews_find_items_all
      .replace('$FolderId$', opts.folderid)
      .replace('$MaxEntriesReturned$', opts.maximum)
    ;
  }

  var first = promise.Promise();
  first
  .then(function() {
    var p = promise.Promise();

    var req = https.request(url, function(res) {
      if (res.statusCode != 200) {
        url.auth = 'username:password';
        var html = '';
        res.on('data', function(data) { html += data; });
        res.on('end', function() {
          p.reject(url.format() + ': ' + res.statusCode + '\n' + html);
        });
        return;
      }
      var data = '';
      res.on('data', function(buffer) {
        data += buffer;
      });
      res.on('end', function() {
        p.resolve(data);
      });
    });
    req.write(xml);
    req.end();
    return p;
  })
  .then(function(data) {
    var p = promise.Promise();

    xml2js.parseString(data, function(err, result) {
      if (err) {
        p.reject(err);
        return;
      }

      // urg, I know
      var items = result['s:Envelope']['s:Body'][0]['m:FindItemResponse'][0]['m:ResponseMessages'][0]['m:FindItemResponseMessage'][0]['m:RootFolder'][0]['t:Items'][0]['t:CalendarItem'];
      items = items.map(function(item) {
        return {
          id: item['t:ItemId'][0]['$']['Id']
        };
      });
      p.resolve(items);
    });

    return p;
  })
  .then(function(items) {
    var p = promise.Promise();

    var xml = self.templates.ews_get_item
      .replace('$FolderId$', opts.folderid)
      .replace('$FieldURI$', fields)
      .replace('$ItemId$', items.map(function(item) {
        return '<t:ItemId Id="' + item.id + '" />'
      }).join('\n'))
    ;

    var req = https.request(url, function(res) {
      if (res.statusCode != 200) {
        url.auth = 'username:password';
        p.reject(url.format() + ': ' + res.statusCode);
        res.on('data', function() {});
        res.on('end', function() {});
        return;
      }
      var data = '';
      res.on('data', function(buffer) {
        data += buffer;
      });
      res.on('end', function() {
        p.resolve(data);
      });
    });
    req.write(xml);
    req.end();
    return p;
  })
  .then(function(data) {
    var p = promise.Promise();

    xml2js.parseString(data, function(err, result) {
      if (err) {
        p.reject(err);
        return;
      }

      var items = result['s:Envelope']['s:Body'][0]['m:GetItemResponse'][0]['m:ResponseMessages'][0]['m:GetItemResponseMessage'];
      //[0]['m:RootFolder'][0]['t:Items'][0]['t:CalendarItem'];
      items = items.map(function(item) {
        item = item['m:Items'][0]['t:CalendarItem'][0];
        return {
          id: item['t:ItemId'][0]['$']['Id'],
          subject: item['t:Subject'][0],
          location: item['t:Location'][0],
          start: item['t:Start'][0],
          end: item['t:End'][0],
          categories: (item['t:Categories'] ? item['t:Categories'] : []).map(function(category) {
            return category['t:String'][0];
          }),
          body: item['t:Body'][0]['_']
        };
      });
      p.resolve(items);
    });

    return p;
  })
  .then(function(items) {
    p.resolve(items);
  },function(err) {
    p.reject(err);
  });
  first.resolve();

  return p;
};
