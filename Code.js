const investingId = "13w1DkZl877zR-EuYFTXTFf5slwtt-pC4"
const sectorsId = "1hxxazSfy6N-SOG64BbzTSwvXy_sk9jQK"

const relevantUsers = ['samiton@princeton.edu', 'zamiton@outlook.com', 'zachary_amiton@brown.edu', 'zack@calyxhealth.com']
const sectorHeader = ['Company', 'Symbol', 'Price', 'Last Updated', 'Subsector', 'Held?', 'Zack Shares', 'Zack Price', 'Zack Change', 'Sophie Shares','Sophie Price', 'Sophie Change']	

const spreadsheet = SpreadsheetApp.getActive()

const homeTab = spreadsheet.getSheetByName("Home")
const keepTabs = ['Home', 'Sector Template', 'Master View']
const sectors = {}
const clearTabs = () => spreadsheet.getSheets().filter(s => !keepTabs.includes(s.getName())).forEach(s => spreadsheet.deleteSheet(s))
const hyperlink = (src, str) => `=HYPERLINK("${src}", "${str}")`

function onOpen(e) {
  const menuItems = [{name: 'Compile Stonks', functionName: 'buildMaster'},
                     {name: 'Update Prices (Sector)', functionName: 'populatePrices'},
                     {name: 'Update Prices (All)', functionName: 'buildPrices'},
                     {name: 'Create Research Docs (Sector)', functionName: 'createDocs'},
                     {name: 'Create Research Docs (All)', functionName: 'buildResearch'}]
  spreadsheet.addMenu('Scripts', menuItems)
}

function createDocs() {
  Logger.log('tbd')
}

function shareUsers(doc, users) {
  users.forEach(u => {
    Drive.Permissions.insert({role: "writer", type: "user", value: u}, doc, {sendNotificationEmails: false})         
  })
}

function populatePrices(currentSector=spreadsheet.getActiveSheet()) {
  const endpointUrl = (sym) => (sym)?`https://query2.finance.yahoo.com/v7/finance/chart/${sym}`:null
  const symbol = sectorHeader.indexOf("SYM")
  const symbolRequests = currentSector.getRange(2, symbol + 1, currentSector.getLastRow() - 1).getValues().flatMap(f => endpointUrl(f[0]))
  
  if(symbolRequests.every(x => !!x)) {  
    const stockResponses = UrlFetchApp.fetchAll(symbolRequests).map(resp => {
      const response = JSON.parse(resp.getContentText())
      if(!response.error) {
        const chart = resp
        if(chart) {
          const result = response.chart.result[0].meta
          return [result.regularMarketPrice, Utilities.formatDate(new Date(1000*(result.regularMarketTime - result.gmtoffset)), "GMT-8", "MM-dd HH:mm:ss")]
        }
        return [null, null]
      }                                                               
    }) 
    currentSector.getRange(2, symbol + 2, currentSector.getLastRow() - 1, 2).setValues(stockResponses)  
  } else {
    SpreadsheetApp.getUi().alert("All listings must have a valid ticker symbol (SYM) in order to populate prices!")    
  }
}

function createDocs(active=null) {
  if(!active) active = spreadsheet.getActiveSheet()
  if(Object.keys(sectors).length == 0) buildSectors()
  if(sectors[active.getName()]) {
    const sectorFolder = DriveApp.getFolderById(sectors[active.getName()])
    const stockRange = active.getRange(2, 1, spreadsheet.getLastRow() - 1)

    const stockFormulas = stockRange.getFormulas().flat()
    const stockNames = stockRange.getValues().flat()
    
    for(let i = 0; i < stockNames.length; i++) {
      let [form, name] = [stockFormulas[i], stockNames[i]]      

      if(name && !form.startsWith('=HYPERLINK')) {   
        Logger.log(form, name)
        const brand = DocumentApp.create(name)
        const brandFile = DriveApp.getFileById(brand.getId())
        brandFile.moveTo(sectorFolder)
        stockFormulas[i] = hyperlink(brandFile.getUrl(), name)  
      }
    }
    stockRange.setValues(stockFormulas.map(s => [s]))
  }
}

function buildResearch() {
  for(let s of spreadsheet.getSheets().filter(s => !keepTabs.includes(s.getName()))) createDocs(s)
}

function buildPrices() {
  for(let s of spreadsheet.getSheets().filter(s => !keepTabs.includes(s.getName()))) populatePrices(s)
}

function buildMaster() {
  const sectorTabs = spreadsheet.getSheets().filter(s => !keepTabs.includes(s.getName()))
  const masterTab = spreadsheet.getSheetByName("Master View") || spreadsheet.getSheetByName("Sector Template").copyTo(spreadsheet).setName("Master View")
  if(masterTab.getLastRow() > 1) masterTab.deleteRows(2, masterTab.getLastRow() - 1)
  
  let stonks = []
  
  for(let st of sectorTabs) {
    const sr = st.getRange(2, 1, st.getLastRow() - 1, st.getLastColumn())
    const formulas = sr.getFormulas()
    const values = sr.getValues()
    stonks.push(formulas.map((f, ridx) => f.map((c, cidx) => c || values[ridx][cidx])).filter(r => !!r[0]))
  }

  if(stonks.length > 0) masterTab.getRange(2, 1, stonks.length + 1, stonks[0][0].length).setValues(stonks.flat()) 
}

function buildSectors() {
  const sectorsFolder = DriveApp.getFolderById(sectorsId)
  const sectorIterator = sectorsFolder.getFolders()
  const existingSectors = new Set()
  while(sectorIterator.hasNext()) {
    const nextSector = sectorIterator.next()
    sectors[nextSector.getName()] = nextSector.getId() 
  }
  
  for(let s of homeTab.getRange(2, 1, homeTab.getLastRow() - 1).getValues().flatMap(r => r[0].trim()).filter(x => !!x)) {
    if(!(s in sectors)) {
       const newFolder = DriveApp.createFolder(s)
       newFolder.moveTo(sectorsFolder)
       sectors[s] = newFolder.getId()
       shareUsers(newFolder.getId(), relevantUsers) 
    } 
    if(!spreadsheet.getSheets().map(sn => sn.getName()).includes(s)) {
      const newSector = spreadsheet.getSheetByName("Sector Template").copyTo(spreadsheet).setName(s)
      spreadsheet.setActiveSheet(newSector)
      spreadsheet.moveActiveSheet(spreadsheet.getNumSheets()-1)
      newSector.showSheet()
    }
  }
  
  relinkSectors()
}

function relinkSectors() {
  const sheets = {}
  const sectorTabs = spreadsheet.getSheets().forEach(s => {
    if(!keepTabs.includes(s.getName())) sheets[s.getName()] = s.getSheetId()
  })
  
  const homeTab = spreadsheet.getSheetByName("Home")
  const sectorRange = homeTab.getRange(2, 1, homeTab.getLastRow() - 1)
  sectorRange.setValues(sectorRange.getDisplayValues().map(x => [`=HYPERLINK("#gid=${sheets[x[0]]}", "${x[0]}")`]))
}
