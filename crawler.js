const xmlParser = require('fast-xml-parser')
const Request = require('request-promise-native')
const moment = require('moment')
const json2xls = require('json2xls')
const fs = require('fs')
const url = require('url')

// Ref from the source https://www.legco.gov.hk/scripts/mtg_vote.js
const yearPages = [
  'https://www.legco.gov.hk/php/detect-votes.php?term=yr16-17&meeting=cm',
  'https://www.legco.gov.hk/php/detect-votes.php?term=yr17-18&meeting=cm',
  'https://www.legco.gov.hk/php/detect-votes.php?term=yr18-19&meeting=cm',
  'https://www.legco.gov.hk/php/detect-votes.php?term=yr16-17&meeting=hc',
  'https://www.legco.gov.hk/php/detect-votes.php?term=yr17-18&meeting=hc',
  'https://www.legco.gov.hk/php/detect-votes.php?term=yr18-19&meeting=hc',
  'https://www.legco.gov.hk/php/detect-votes.php?term=yr16-17&meeting=fc',
  'https://www.legco.gov.hk/php/detect-votes.php?term=yr17-18&meeting=fc',
  'https://www.legco.gov.hk/php/detect-votes.php?term=yr18-19&meeting=fc',
  'https://www.legco.gov.hk/php/detect-votes.php?term=yr16-17&meeting=esc',
  'https://www.legco.gov.hk/php/detect-votes.php?term=yr17-18&meeting=esc',
  'https://www.legco.gov.hk/php/detect-votes.php?term=yr18-19&meeting=esc',
  'https://www.legco.gov.hk/php/detect-votes.php?term=yr16-17&meeting=pwsc',
  'https://www.legco.gov.hk/php/detect-votes.php?term=yr17-18&meeting=pwsc',
  'https://www.legco.gov.hk/php/detect-votes.php?term=yr18-19&meeting=pwsc']

function getXmlFolder(meeting){
  switch(meeting)
  {
    case 'cm': return 'counmtg/voting'; break;
    case 'hc': return 'hc/voting'; break;
    case 'fc': return 'fc/fc/results'; break;
    case 'esc': return 'fc/esc/results'; break;
    case 'pwsc': return 'fc/pwsc/results'; break;
    default: return 'counmtg/voting'; break;
  }
}


let nameIndex = {}
let voteData = {}
let excelJson = []

async function getXmlFiles(){
  console.log('******* Load data *******')
  for (let i=0; i<yearPages.length; i++ ){
    let uri = yearPages[i]
    const myURL = url.parse(uri, true)
    let term = myURL.query['term']
    let meeting = myURL.query['meeting']
    let xmlFolder = getXmlFolder(meeting)
    console.log([term, meeting])

    let fileList = await Request.get(uri)
    let xmlFiles = fileList.split(',').filter(filename => filename.indexOf('.xml') !== -1)
    await loadXml(term, xmlFolder, meeting, xmlFiles)
  }
  console.log('******* Gen excel *******')
  transformData()
  exportExcel()
}

function transformData(){
  Object.keys(voteData).forEach( (key, idx) => {
    let rowData = {
      VoteDate: voteData[key].date.toDate(),
      VoteTitleTC: voteData[key].titleTC,
      VoteTitleEN: voteData[key].titleEN
    }
    Object.assign(rowData, nameIndex);

    let voteResult = voteData[key].voteResult
    Object.keys(voteResult).forEach( (name) => {
      rowData[name] = voteResult[name]
    })

    excelJson.push(rowData)
  })
}

function exportExcel(){
  excelJson.sort(function(a, b){
    return a.VoteDate - b.VoteDate
  })
  var xls = json2xls(excelJson);
  fs.writeFileSync('vote-result.xlsx', xls, 'binary');
}

async function loadXml(term, xmlFolder, meeting, files){
  for (let i=0; i<files.length; i++ ){
    let xmlLink = `https://www.legco.gov.hk/${term}/chinese/${xmlFolder}/${meeting}_vote_${files[i]}`
    await parseXML(xmlLink)
  }
  return 
}

async function parseXML(fileUrl){
  let parseOptions = {
    ignoreAttributes: false
  }

  try{
    let xmlData = await Request.get(fileUrl)

    if(xmlParser.validate(xmlData)){
      let jsonData = xmlParser.parse(xmlData, parseOptions)['legcohk-vote'].meeting.vote
      if(jsonData.length === undefined) { // handle the meeting only has 1 vote
        let temp = jsonData
        jsonData = [temp]
      }
      console.log(`${fileUrl} vote count:${jsonData.length}`)

      jsonData.forEach( voteResult => {
        let voteDatetime = moment(`${voteResult['vote-date']} ${voteResult['vote-time']}`, 'DD/MM/YYYY HH:mm:ss')
        // let voteDatetime = moment(`${voteResult['vote-date']}`, 'DD/MM/YYYY')
        let voteTitle = `${voteResult['motion-ch']} ${voteResult['motion-en']}`
        let individualVotes = voteResult['individual-votes']['member']
        // console.log(voteDatetime)
        // console.log(voteTitle)
        // console.log(individualVotes)

        voteData[encodeURI(voteTitle)] = {
          titleEN: voteResult['motion-en'],
          titleTC: voteResult['motion-ch'],
          date: voteDatetime,
          voteResult: {}
        }

        individualVotes.forEach(function(item){
          // let name = `${item['@_name-ch']} ${item['@_name-en']}`
          let name = `${item['@_name-ch']}` // some councillors change their english name, so use chinese name as key
          nameIndex[name] = '-' // build up full councillor name list, value is default vote result in excel for the DQ councillor
          voteData[encodeURI(voteTitle)].voteResult[name] = item['vote']
        })
      })
      return true
    } else {
      throw Error('XML melform')
    }
  } catch(err){
    return err
  }
}

getXmlFiles()