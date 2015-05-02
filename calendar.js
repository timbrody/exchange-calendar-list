var
  fs = require('fs')
  , promise = require('node-promise')
  , EventEmitter = require('events').EventEmitter
  , util = require('util')
  , https = require('https')
  , xpath = require('xpath')
  , dom = require('xmldom').DOMParser
  , soap = require('soap')
  , path = require('path')
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
      "calendar:Location",
      "calendar:Start",
      "calendar:End",
      "item:Categories"
    ];

  var fields = '';
  opts.fields.forEach(function(field) {
    fields += '<t:FieldURI FieldURI="' + field + '"/>\n';
  });
  
  var client = null;
 
var first = promise.Promise();
first
.then(function(){
	var p = promise.Promise();
			
	var endpoint = url.href;
	var wsdl = path.join(__dirname, 'Services.wsdl');
	
	console.log(endpoint);
	    
	soap.createClient(wsdl, {}, function(err, clnt) {
	  if (err) {
	    return p.reject(err);
	  }
	  
	  client = clnt;
	  if (!client) {
	    return p.reject('Could not create client');
	  }
      
	  client.setSecurity(new soap.BasicAuthSecurity(opts.username, opts.password));
	  p.resolve(client);
	  return p;
	}, endpoint);
	
	return p;
})
.then(function(client){
	var p = promise.Promise();
	
	if (!opts.folderid) {
		var xml = self.templates.ews_get_folder;
		
		client.GetFolder(xml, function(err, result, body) {
			if (err) {
			  return p.reject(err.message);
			}
			
			var doc = new dom().parseFromString(body);
			var select = xpath.useNamespaces({"t": "http://schemas.microsoft.com/exchange/services/2006/types"});
			var nodes = select("//t:FolderId", doc);
			if (nodes.length > 0) {
			   p.resolve(nodes[0].getAttribute('Id'));
			}
			return p;
		});

    return p;
	}
	return opts.folderid;
})
.then(function(folderid) {
	opts.folderid = folderid;
	var xml;
	if (opts.start) {
	  var start = new Date(opts.start);
	  var end = new Date(opts.end);
	
	  start.toString = end.toString = function() {
	    return this.toISOString().replace(/\..+/, '') + '-00:00';
	  };
      
	  xml = self.templates.ews_find_items
	    .replace('$FolderId$', folderid)
	    .replace('$StartDate$', start)
	    .replace('$EndDate$', end)
	    .replace('$FieldURI$', fields)
	    .replace('$MaxEntriesReturned$', opts.maximum);
	}
	else {
	  xml = self.templates.ews_find_items_all
	    .replace('$FolderId$', folderid)
	    .replace('$FieldURI$', fields)
	    .replace('$MaxEntriesReturned$', opts.maximum);
	}
	
	return xml;
})
.then(function(xml) {
    var p = promise.Promise();
    client.FindItem(xml, function(err, result, body) {
			if (err) {
			  return p.reject(err);
			}
			
			p.resolve(body);
			return p;
		});
    
    return p;
})
.then(function(data) {
	var p = promise.Promise();
	    
	var doc = new dom().parseFromString(data);
	var select = xpath.useNamespaces({"t": "http://schemas.microsoft.com/exchange/services/2006/types"});
	var nodes = select("//t:CalendarItem/t:ItemId", doc);
	
	items = nodes.map(function(item) {
	   return {
	     itemId: item.getAttribute('Id'),
	     changeKey: item.getAttribute('ChangeKey')
	  };
	});
	p.resolve(items);

  return p;
})
.then(function(items) {
    var p = promise.Promise();

    var xml = self.templates.ews_get_item
      .replace('$FolderId$', opts.folderid)
      .replace('$FieldURI$', fields)
      .replace('$ItemId$', items.map(function(item) {
        return '<t:ItemId Id="' + item.itemId + '" ChangeKey="' + item.changeKey + '"/>'
      }).join('\n'));
    
	  client.GetItem(xml, function(err, result, body) {
				if (err) {
					return p.reject(err);
				}
				
				p.resolve(body);
				return p;
	  });
    return p;
})
.then(function(data) {
    var p = promise.Promise();
    
    var doc = new dom().parseFromString(data);
    var select = xpath.useNamespaces({"t": "http://schemas.microsoft.com/exchange/services/2006/types"});
    var nodes = select("//t:CalendarItem", doc);

    items = nodes.map(function(item) {
				return {
					itemId: select('./t:ItemId', item)[0].getAttribute('Id'),
					changeKey: select('./t:ItemId', item)[0].getAttribute('ChangeKey'),
					type: firstData(select, './t:CalendarItemType', item),
					subject: firstData(select, './t:Subject', item),
					location: firstData(select, './t:Location', item),
					start: firstData(select, './t:Start', item),
					end: firstData(select, './t:End', item),
					allDay: firstData(select, './t:IsAllDayEvent', item) == 'true' ? true : false,
					body: firstData(select, './t:Body', item),
					organizer: person(select, './t:Organizer', item)[0],
					requiredAttendees: person(select, './t:RequiredAttendees/t:Attendee', item),
					optionalAttendees: person(select, './t:OptionalAttendees/t:Attendee', item)
				};
    });
   
    return items;
})
.then(function(items) {
    p.resolve(items);
  },function(err) {
    p.reject(err);
});
  
	first.resolve();
  return p;
};

function person(select, query, node)
{
  if (node == null) {
				return null;
	}
	
	var ret = [];
  var items = select(query, node);
	for(var i = 0; i < items.length; i++)
	{
				var retItem = {};
				var mb = select('./t:Mailbox/t:Name', items[i])[0];
				if (mb && mb.firstChild) {
				    retItem.name = mb.firstChild.data;
				}
				mb = select('./t:Mailbox/t:EmailAddress', items[i])[0];
				if (mb && mb.firstChild) {
				    retItem.email = mb.firstChild.data;
				}
				ret.push(retItem);
	}
	return ret;
}

function firstData(select, query, node)
{
  var item = select(query, node)[0];
  return item && item.firstChild ? item.firstChild.data : null;
}
