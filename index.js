var path = require('path')
  , moment = require('moment')
  , crypto = require('crypto')
  , xml2js = require('xml2js')
  ;
var Q = require('q');
var soap = require('soap-ntlm-2');

var __ = function() {};

/**
 * Initialize a soap client to communicate with the EWS server
 * @param settings
 * @param settings.username
 * @param settings.password
 * @param settings.url - host url for ews server. EX: mail.huntingtonlearningcenter.com
 *
 * @return {promise - exchanger client}
 */
__.prototype.initialize = function(settings) {
    var instance = this;

    instance.username = settings.username;
    instance.security = {
        basic: new soap.BasicAuthSecurity(settings.username, settings.password),
        ntlm: new soap.NtlmSecurity(settings.username, settings.password),
        ws: new soap.WSSecurity(settings.username, settings.password)
    };

    // TODO: Handle different locations of where the asmx lives.
    var endpoint = 'https://' + path.join(settings.url, 'EWS/Exchange.asmx');
    var url = path.join(__dirname, 'Services.wsdl');

    return Q.nfcall(soap.createClient, url, {
        endpoint: endpoint
    }).then(
        function(client) {
          instance.client = client;
          instance.setSecurity('ntlm');
          return client;
        }
    );
};

/**
 * Sets the security to be used. Default basic security
 * @param {string} type - in [basic, ntlm, or ws]
 */
__.prototype.setSecurity = function(type) {
    type = type || 'basic';
    if (Object.keys(this.security).indexOf(type) < 0) {
        return;
    }
    this.client.setSecurity(this.security[type]);
};

/**
 * Gets all emails from
 * @param {string} folderName - default inbox
 * @param {integer} limit - default 10
 * @return {promise - array email objects}
 */
__.prototype.getEmails = function(folderName, limit) {
    folderName = folderName || 'inbox';
    limit = limit || 10;

    var soapRequest =
        '<tns:FindItem Traversal="Shallow" xmlns:tns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
        '<tns:ItemShape>' +
            '<t:BaseShape>IdOnly</t:BaseShape>' +
            '<t:AdditionalProperties>' +
            '<t:FieldURI FieldURI="item:ItemId"></t:FieldURI>' +
            // '<t:FieldURI FieldURI="item:ConversationId"></t:FieldURI>' +
            // '<t:FieldURI FieldURI="message:ReplyTo"></t:FieldURI>' +
            // '<t:FieldURI FieldURI="message:ToRecipients"></t:FieldURI>' +
            // '<t:FieldURI FieldURI="message:CcRecipients"></t:FieldURI>' +
            // '<t:FieldURI FieldURI="message:BccRecipients"></t:FieldURI>' +
            '<t:FieldURI FieldURI="item:DateTimeCreated"></t:FieldURI>' +
            '<t:FieldURI FieldURI="item:DateTimeSent"></t:FieldURI>' +
            '<t:FieldURI FieldURI="item:HasAttachments"></t:FieldURI>' +
            '<t:FieldURI FieldURI="item:Size"></t:FieldURI>' +
            '<t:FieldURI FieldURI="message:From"></t:FieldURI>' +
            '<t:FieldURI FieldURI="message:IsRead"></t:FieldURI>' +
            '<t:FieldURI FieldURI="item:Importance"></t:FieldURI>' +
            '<t:FieldURI FieldURI="item:Subject"></t:FieldURI>' +
            '<t:FieldURI FieldURI="item:DateTimeReceived"></t:FieldURI>' +
            '</t:AdditionalProperties>' +
        '</tns:ItemShape>' +
        '<tns:IndexedPageItemView BasePoint="Beginning" Offset="0" MaxEntriesReturned="10"></tns:IndexedPageItemView>' +
        '<tns:ParentFolderIds>' +
            '<t:DistinguishedFolderId Id="inbox"></t:DistinguishedFolderId>' +
        '</tns:ParentFolderIds>' +
        '</tns:FindItem>';

    return Q.nfcall(this.client.FindItem, soapRequest).spread(
        function(result, body) {
            var parser = new xml2js.Parser();

            return Q.nfcall(parser.parseString, body).then(
                function(result) {
                    var responseCode = result['s:Body']['m:FindItemResponse']['m:ResponseMessages']['m:FindItemResponseMessage']['m:ResponseCode'];

                    if (responseCode !== 'NoError') {
                        return callback(new Error(responseCode));
                    }

                    var rootFolder = result['s:Body']['m:FindItemResponse']['m:ResponseMessages']['m:FindItemResponseMessage']['m:RootFolder'];

                    var emails = [];
                    rootFolder['t:Items']['t:Message'].forEach(function(item, idx) {
                        var md5hasher = crypto.createHash('md5');
                        md5hasher.update(item['t:Subject'] + item['t:DateTimeSent']);
                        var hash = md5hasher.digest('hex');

                        var itemId = {
                        id: item['t:ItemId']['@'].Id,
                        changeKey: item['t:ItemId']['@'].ChangeKey
                        };

                        var dateTimeReceived = item['t:DateTimeReceived'];

                        emails.push({
                            id: itemId.id + '|' + itemId.changeKey,
                            hash: hash,
                            subject: item['t:Subject'],
                            dateTimeReceived: moment(dateTimeReceived).format("MM/DD/YYYY, h:mm:ss A"),
                            size: item['t:Size'],
                            importance: item['t:Importance'],
                            hasAttachments: (item['t:HasAttachments'] === 'true'),
                            from: item['t:From']['t:Mailbox']['t:Name'],
                            isRead: (item['t:IsRead'] === 'true'),
                            meta: {
                                itemId: itemId
                            }
                        });
                    });
                    return emails;
                }
            );
        }
    );
}

/**
 * Resolve a name
 * @param {string} name - name to resolve
 * @return {promise - array email objects}
 */
__.prototype.resolveNames = function(name) {
  if (!name) {
    return callback(new Error('No name provided.'));
  }

  var soapRequest =
    '<tns:ResolveNames xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" ReturnFullContactData="false">' +
      '<tns:UnresolvedEntry>' + name + '</tns:UnresolvedEntry>' +
    '</tns:ResolveNames>';

  return Q.nfcall(this.client.ResolveNames, soapRequest).spread(
    function(result, body) {
      var parser = new xml2js.Parser();

      return Q.nfcall(parser.parseString, body).then(
        function(result) {
          var responseCode = result['s:Body']['m:ResolveNamesResponse']['m:ResponseMessages']['m:ResolveNamesResponseMessage']['m:ResponseCode'];
          if (responseCode !== 'NoError') {
            throw new Error(responseCode);
          }

          var rootFolder = result['s:Body']['m:ResolveNamesResponse']['m:ResponseMessages']['m:ResolveNamesResponseMessage']['m:ResolutionSet'];
          var contacts = [], iterableResult = [];

          if(rootFolder['@'].TotalItemsInView == 1) {
            iterableResult.push(rootFolder['t:Resolution']);
          }
          else {
            iterableResult = rootFolder['t:Resolution'];
          }
          iterableResult.forEach(function (item, idx) {
            contacts.push({
              name: item['t:Mailbox']['t:Name'],
              email: item['t:Mailbox']['t:EmailAddress'],
            });
          });

          return contacts;
        }
      )
    }
  ).fail(function(result) {
    if (typeof result.response !== 'undefined' && result.response.statusCode == 401) {
      var error = new Error('Unauthorized');
      error.code = 401;
      throw error;
    }
    else if (result.code == 'ENOTFOUND') {
      var error = new Error('Not Found');
      error.code = 404;
      throw error;
    }
  });
}

/**
 * Checks if the account can be successfully accessed
 * Currently uses resolveNames until a better function is coded.
 * @returns {promise resolves if login successful}
 */
__.prototype.checkLogin = __.prototype.resolveNames;


/**
 * Other preconstructed EWS soap requests
 */

// __.prototype.getEmail = function(itemId, callback) {

//   if ((!itemId['id'] || !itemId['changeKey']) && itemId.indexOf('|') > 0) {
//     var s = itemId.split('|');

//     itemId = {
//       id: itemId.split('|')[0],
//       changeKey: itemId.split('|')[1]
//     };
//   }

//   if (!itemId.id || !itemId.changeKey) {
//     return callback(new Error('Id is not correct.'));
//   }

//   var soapRequest =
//     '<tns:GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
//       '<tns:ItemShape>' +
//         '<t:BaseShape>Default</t:BaseShape>' +
//         '<t:IncludeMimeContent>true</t:IncludeMimeContent>' +
//       '</tns:ItemShape>' +
//       '<tns:ItemIds>' +
//         '<t:ItemId Id="' + itemId.id + '" ChangeKey="' + itemId.changeKey + '" />' +
//       '</tns:ItemIds>' +
//     '</tns:GetItem>';

//   this.client.GetItem(soapRequest, function(err, result, body) {
//     if (err) {
//       return callback(err);
//     }

//     var parser = new xml2js.Parser();

//     parser.parseString(body, function(err, result) {
//       var responseCode = result['s:Body']['m:GetItemResponse']['m:ResponseMessages']['m:GetItemResponseMessage']['m:ResponseCode'];

//       if (responseCode !== 'NoError') {
//         return callback(new Error(responseCode));
//       }

//       var item = result['s:Body']['m:GetItemResponse']['m:ResponseMessages']['m:GetItemResponseMessage']['m:Items']['t:Message'];

//       var itemId = {
//         id: item['t:ItemId']['@'].Id,
//         changeKey: item['t:ItemId']['@'].ChangeKey
//       };

//       function handleMailbox(mailbox) {
//         var mailboxes = [];

//         if (!mailbox || !mailbox['t:Mailbox']) {
//           return mailboxes;
//         }
//         mailbox = mailbox['t:Mailbox'];

//         function getMailboxObj(mailboxItem) {
//           return {
//             name: mailboxItem['t:Name'],
//             emailAddress: mailboxItem['t:EmailAddress']
//           };
//         }

//         if (mailbox instanceof Array) {
//           mailbox.forEach(function(m, idx) {
//             mailboxes.push(getMailboxObj(m));
//           })
//         } else {
//           mailboxes.push(getMailboxObj(mailbox));
//         }

//         return mailboxes;
//       }

//       var toRecipients = handleMailbox(item['t:ToRecipients']);
//       var ccRecipients = handleMailbox(item['t:CcRecipients']);
//       var from = handleMailbox(item['t:From']);

//       var email = {
//         id: itemId.id + '|' + itemId.changeKey,
//         subject: item['t:Subject'],
//         bodyType: item['t:Body']['@']['t:BodyType'],
//         body: item['t:Body']['#'],
//         size: item['t:Size'],
//         dateTimeSent: item['t:DateTimeSent'],
//         dateTimeCreated: item['t:DateTimeCreated'],
//         toRecipients: toRecipients,
//         ccRecipients: ccRecipients,
//         from: from,
//         isRead: (item['t:IsRead'] == 'true') ? true : false,
//         meta: {
//           itemId: itemId
//         }
//       };

//       callback(null, email);
//     });
//   });
// }


// __.prototype.getFolders = function(id, callback) {
//   if (typeof(id) == 'function') {
//     callback = id;
//     id = 'inbox';
//   }

//   var soapRequest =
//     '<tns:FindFolder xmlns:tns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
//         '<tns:FolderShape>' +
//           '<t:BaseShape>Default</t:BaseShape>' +
//         '</tns:FolderShape>' +
//         '<tns:ParentFolderIds>' +
//           '<t:DistinguishedFolderId Id="inbox"></t:DistinguishedFolderId>' +
//         '</tns:ParentFolderIds>' +
//       '</tns:FindFolder>';

//   this.client.FindFolder(soapRequest, function(err, result) {
//     if (err) {
//       callback(err)
//     }

//     if (result.ResponseMessages.FindFolderResponseMessage.ResponseCode == 'NoError') {
//       var rootFolder = result.ResponseMessages.FindFolderResponseMessage.RootFolder;

//       rootFolder.Folders.Folder.forEach(function(folder) {
//         // console.log(folder);
//       });

//       callback(null, {});
//     }
//   });
// }

module.exports = new __();
