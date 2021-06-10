const PizZip = require('pizzip')
const Docxtemplater = require('docxtemplater')
const fs = require('fs')
const path = require('path')

function replaceErrors(key, value) {
  if (value instanceof Error) {
    return Object.getOwnPropertyNames(value).reduce(function (error, key) {
      error[key] = value[key]
      return error
    }, {})
  }
  return value
}

function errorHandler(error) {
  console.log(JSON.stringify({ error: error }, replaceErrors))

  if (error.properties && error.properties.errors instanceof Array) {
    const errorMessages = error.properties.errors.map(function (error) {
      return error.properties.explanation
    }).join("\n")
    console.log('errorMessages', errorMessages)
  }
  throw error
}

let content = fs.readFileSync(path.resolve(__dirname, 'test/input.docx'), 'binary')

let zip = new PizZip(content)
let doc
try {
  doc = new Docxtemplater(zip)
} catch (error) {
  errorHandler(error)
}

doc.setData({
  first_name: 'John',
  last_name: 'Doe',
  phone: '0652455478',
  description: 'New Website'
})

try {
  doc.render()
}
catch (error) {
  errorHandler(error)
}

let buf = doc.getZip().generate({ type: 'nodebuffer' })

fs.writeFileSync(path.resolve(__dirname, 'test/output.docx'), buf)

