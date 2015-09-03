var async = require('async')
var request = require('request')
var Parser = require('xml2js').Parser
var http = require('http')
var querystring = require('querystring')
var _ = require('lodash')
var GoogleAuth = require('google-auth-library')
var Q = require('bluebird')

var GOOGLE_FEED_URL = 'https://spreadsheets.google.com/feeds/'
var GOOGLE_AUTH_SCOPE = ['https://spreadsheets.google.com/feeds']

// The main class that represents a single sheet
// this is the main module.exports
function GoogleSpreadsheet (ssKey, authId, options) {
  var self = this
  var googleAuth = null
  var visibility = 'public'
  var projection = 'values'

  var authMode = 'anonymous'

  var authClient = new GoogleAuth()
  var jwtClient

  options = options || {}

  var xmlParser = new Parser({
    // options carried over from older version of xml2js
    // might want to update how the code works, but for now this is fine
    explicitArray: false,
    explicitRoot: false
  })

  if (!ssKey) {
    throw new Error('Spreadsheet key not provided.')
  }

  // authId may be null
  setAuthAndDependencies(authId)

  // Authentication Methods
  this.setAuthToken = function (authId) {
    if (authMode === 'anonymous') authMode = 'token'
    setAuthAndDependencies(authId)
  }

  // deprecated username/password login method
  // leaving it here to help notify users why it doesn't work
  this.setAuth = function (username, password, cb) {
    return Q.resolve().then(function () {
      throw new Error('Google has officially deprecated ClientLogin. ' +
        'Please upgrade this module and see the readme for more instructions')
    }).nodeify(cb)
  }

  this.useServiceAccountAuth = function (creds, cb) {
    if (typeof creds === 'string') creds = require(creds)

    jwtClient = new authClient.JWT(creds.client_email, null, creds.private_key, GOOGLE_AUTH_SCOPE, null)

    return Q.resolve().then(function () {
      return renewJwtAuth()
    }).nodeify(cb)
  }

  function renewJwtAuth (cb) {
    authMode = 'jwt'

    return new Q(function (resolve, reject) {
      jwtClient.authorize(function (err, token) {
        if (err) return reject(err)
        self.setAuthToken({
          type: token.token_type,
          value: token.access_token,
          expires: token.expiry_date
        })
        resolve()
      })
    }).nodeify(cb)
  }

  function setAuthAndDependencies (auth) {
    googleAuth = auth
    if (!options.visibility) {
      visibility = googleAuth ? 'private' : 'public'
    }
    if (!options.projection) {
      projection = googleAuth ? 'full' : 'values'
    }
  }

  // This method is used internally to make all requests
  this.makeFeedRequest = function (urlParams, method, queryOrData, cb) {
    var url
    var headers = {}

    if (typeof (urlParams) === 'string') {
      // used for edit / delete requests
      url = urlParams
    } else if (Array.isArray(urlParams)) {
      // used for get and post requets
      urlParams.push(visibility, projection)
      url = GOOGLE_FEED_URL + urlParams.join('/')
    }

    return new Q(function (resolve, reject) {
      async.series({
        auth: function (step) {
          if (authMode !== 'jwt') return step()

          // check if jwt token is expired
          if (googleAuth.expires > +new Date()) return step()
          renewJwtAuth(step)
        },
        request: function (result, step) {
          if (googleAuth) {
            if (googleAuth.type === 'Bearer') {
              headers['Authorization'] = 'Bearer ' + googleAuth.value
            } else {
              headers['Authorization'] = 'GoogleLogin auth=' + googleAuth
            }
          }

          if (method === 'POST' || method === 'PUT') {
            headers['content-type'] = 'application/atom+xml'
          }

          if (method === 'GET' && queryOrData) {
            url += '?' + querystring.stringify(queryOrData)
          }

          request({
            url: url,
            method: method,
            headers: headers,
            body: method === 'POST' || method === 'PUT' ? queryOrData : null
          }, function (err, response, body) {
            if (err) {
              return reject(err)
            } else if (response.statusCode === 401) {
              return reject(new Error('Invalid authorization key.'))
            } else if (response.statusCode >= 400) {
              return reject(new Error('HTTP error ' + response.statusCode + ': ' + http.STATUS_CODES[response.statusCode]) + ' ' + JSON.stringify(body))
            } else if (response.statusCode === 200 && response.headers['content-type'].indexOf('text/html') >= 0) {
              return reject(new Error('Sheet is private. Use authentication or make public. (see https://github.com/theoephraim/node-google-spreadsheet#a-note-on-authentication for details)'))
            }

            if (body) {
              xmlParser.parseString(body, function (err, result) {
                if (err) return reject(err)
                resolve([result, body])
              })
            } else {
              if (err) reject(err)
              else resolve([true])
            }
          })
        }
      })
    })
  }

  // public API methods
  this.getInfo = function (cb) {
    return self.makeFeedRequest(['worksheets', ssKey], 'GET', null).spread(function (data, xml) {
      if (data === true) {
        throw new Error('No response to getInfo call')
      }
      var ss_data = {
        title: data.title['_'],
        updated: data.updated,
        author: data.author,
        worksheets: []
      }
      var worksheets = forceArray(data.entry)
      worksheets.forEach(function (ws_data) {
        ss_data.worksheets.push(new SpreadsheetWorksheet(self, ws_data))
      })
      return ss_data
    }).nodeify(cb)
  }

  // NOTE: worksheet IDs start at 1

  this.getRows = function (worksheetId, opts, cb) {
    // the first row is used as titles/keys and is not included

    // opts is optional
    if (!cb && !opts) {
      opts = {}
    }
    if (typeof opts === 'function') {
      cb = opts
      opts = {}
    }

    var query = {}

    if (opts.start) query['start-index'] = opts.start
    if (opts.num) query['max-results'] = opts.num
    if (opts.orderby) query['orderby'] = opts.orderby
    if (opts.reverse) query['reverse'] = opts.reverse
    if (opts.query) query['sq'] = opts.query

    return self.makeFeedRequest(['list', ssKey, worksheetId], 'GET', query).spread(function (data, xml) {
      if (data === true) {
        throw new Error('No response to getRows call')
      }

      // gets the raw xml for each entry -- this is passed to the row object so we can do updates on it later
      var entries_xml = xml.match(/<entry[^>]*>([\s\S]*?)<\/entry>/g)
      var entries = forceArray(data.entry)

      return entries.map(function (data, i) {
        return new SpreadsheetRow(self, data, entries_xml[i])
      })
    }).nodeify(cb)
  }

  this.addRow = function (worksheetId, data, cb) {
    var dataXml = '<entry xmlns="http://www.w3.org/2005/Atom" xmlns:gsx="http://schemas.google.com/spreadsheets/2006/extended">\n'

    Object.keys(data).forEach(function (key) {
      if (key !== 'id' && key !== 'title' && key !== 'content' && key !== '_links') {
        dataXml += '<gsx:' + xmlSafeColumnName(key) + '>' + xmlSafeValue(data[key]) + '</gsx:' + xmlSafeColumnName(key) + '>\n'
      }
    })

    dataXml += '</entry>'

    return self.makeFeedRequest(['list', ssKey, worksheetId], 'POST', dataXml).nodeify(cb)
  }

  this.getCells = function (worksheetId, opts, cb) {
    // opts is optional
    if (typeof (opts) === 'function') {
      cb = opts
      opts = {}
    }

    // Supported options are:
    // min-row, max-row, min-col, max-col, return-empty
    var query = _.assign({}, opts)

    return self.makeFeedRequest(['cells', ssKey, worksheetId], 'GET', query).then(function (data, xml) {
      if (data === true) {
        return cb(new Error('No response to getCells call'))
      }

      var cells = []
      var entries = forceArray(data['entry'])

      entries.forEach(function (cell_data) {
        cells.push(new SpreadsheetCell(self, worksheetId, cell_data))
      })

      return cells
    }).nodeify(cb)
  }
}

// Classes
function SpreadsheetWorksheet (spreadsheet, data) {
  var self = this

  self.id = data.id.substring(data.id.lastIndexOf('/') + 1)
  self.title = data.title['_']
  self.rowCount = data['gs:row_count']
  self.colCount = data['gs:col_count']

  this.getRows = function (opts, cb) {
    return spreadsheet.getRows(self.id, opts).nodeify(cb)
  }
  this.getCells = function (opts, cb) {
    return spreadsheet.getCells(self.id, opts).nodeify(cb)
  }
  this.addRow = function (data, cb) {
    return spreadsheet.addRow(self.id, data).nodeify(cb)
  }
}

function SpreadsheetRow (spreadsheet, data, xml) {
  var self = this
  self['_xml'] = xml
  Object.keys(data).forEach(function (key) {
    var val = data[key]
    if (key.substring(0, 4) === 'gsx:') {
      if (typeof val === 'object' && Object.keys(val).length === 0) {
        val = null
      }
      if (key === 'gsx:') {
        self[key.substring(0, 3)] = val
      } else {
        self[key.substring(4)] = val
      }
    } else {
      if (key === 'id') {
        self[key] = val
      } else if (val['_']) {
        self[key] = val['_']
      } else if (key === 'link') {
        self['_links'] = []
        val = forceArray(val)
        val.forEach(function (link) {
          self._links[link['$']['rel']] = link['$']['href']
        })
      }
    }
  }, this)

  self.save = function (cb) {
    /*
    API for edits is very strict with the XML it accepts
    So we just do a find replace on the original XML.
    It's dumb, but I couldnt get any JSON->XML conversion to work reliably
    */

    var dataXml = self['_xml']
    // probably should make this part more robust?
    dataXml = dataXml.replace('<entry>', "<entry xmlns='http://www.w3.org/2005/Atom' xmlns:gsx='http://schemas.google.com/spreadsheets/2006/extended'>")
    Object.keys(self).forEach(function (key) {
      if (key.substr(0, 1) !== '_' && typeof (self[key] === 'string')) {
        dataXml = dataXml.replace(new RegExp('<gsx:' + xmlSafeColumnName(key) + '>([\\s\\S]*?)</gsx:' + xmlSafeColumnName(key) + '>'), '<gsx:' + xmlSafeColumnName(key) + '>' + xmlSafeValue(self[key]) + '</gsx:' + xmlSafeColumnName(key) + '>')
      }
    })
    spreadsheet.makeFeedRequest(self._links.edit, 'PUT', dataXml, cb)
  }
  self.del = function (cb) {
    return spreadsheet.makeFeedRequest(self._links.edit, 'DELETE', null).nodeify(cb)
  }
}

var SpreadsheetCell = function (spreadsheet, worksheetId, data) {
  var self = this
  var cell = data['gs:cell']

  self.id = data['id']
  self.row = parseInt(cell.$.row, 10)
  self.col = parseInt(cell.$.col, 10)
  self.value = cell._
  self.numericValue = cell.$.numericValue

  self['_links'] = []
  var links = forceArray(data.link)
  links.forEach(function (link) {
    self._links[link['$']['rel']] = link['$']['href']
  })

  self.setValue = function (newValue, cb) {
    self.value = newValue
    self.save(cb)
  }

  self.save = function (cb) {
    var newValue = xmlSafeValue(self.value)
    var editId = 'https://spreadsheets.google.com/feeds/cells/key/worksheetId/private/full/R' + self.row + 'C' + self.col
    var dataXml =
    '<entry><id>' + editId + '</id>' +
      '<link rel="edit" type="application/atom+xml" href="' + editId + '"/>' +
      '<gs:cell row="' + self.row + '" col="' + self.col + '" inputValue="' + newValue + '"/></entry>'

    dataXml = dataXml.replace('<entry>', "<entry xmlns='http://www.w3.org/2005/Atom' xmlns:gs='http://schemas.google.com/spreadsheets/2006'>")

    spreadsheet.makeFeedRequest(self._links.edit, 'PUT', dataXml, cb)
  }

  self.del = function (cb) {
    self.setValue('', cb)
  }
}

module.exports = GoogleSpreadsheet

GoogleSpreadsheet.SpreadsheetCell = SpreadsheetCell
GoogleSpreadsheet.SpreadsheetRow = SpreadsheetRow
GoogleSpreadsheet.SpreadsheetWorksheet = SpreadsheetWorksheet

// utils
var forceArray = function (val) {
  if (Array.isArray(val)) return val
  if (!val) return []
  return [ val ]
}
var xmlSafeValue = function (val) {
  if (val === null) return ''
  return String(val).replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}
var xmlSafeColumnName = function (val) {
  if (!val) return ''
  return String(val).replace(/[\s_]+/g, '')
    .toLowerCase()
}
