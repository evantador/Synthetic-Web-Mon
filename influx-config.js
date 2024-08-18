
const {InfluxDB, Point} = require('@influxdata/influxdb-client')
const url = 'http://localhost:8086';  // Replace with your InfluxDB URL
const token = '1LZ1AQxMyUA8cmecHaAlbhxJtm56q1qiqqZ415tOn8cn34SfqcMVdcUSFzp6fjrGdbao3zFxOLM_RmDMYTsFSw==';      // Replace with your token
const client = new InfluxDB({url, token})

let org = `Bajaj`
let bucket = `Bajaj Monitoring`

let writeClient = client.getWriteApi(org, bucket, 'ns')

queryClient = client.getQueryApi(org)
fluxQuery = `CREATE DATABASE bajaj_finance_data`

queryClient.queryRows(fluxQuery, {
  next: (row, tableMeta) => {
    const tableObject = tableMeta.toObject(row)
    console.log(tableObject)
  },
  error: (error) => {
    console.error('\nError', error)
  },
  complete: () => {
    console.log('\nSuccess')
  },
})