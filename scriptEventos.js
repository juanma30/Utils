const fs = require('fs')
const path = require('path')
const xlsx = require('exceljs')

const info = [{
    "path": "10.csv",
    "Evento": "The Issue",
    "Mes": "Agosto",
    "Fecha": "03-Agosto-2022",
    "Semana": 20,
    "Hora": "10:00 am",
    "Impartido por": "GS1"
},
{
    "path": "10.csv",
    "Evento": "Control de inventarios a través de RFID: mitos y realidades",
    "Mes": "Agosto",
    "Fecha": "04-Agosto-2022",
    "Semana": 20,
    "Hora": "10:00 am",
    "Impartido por": "Profesionales En Inventarios"
},
{
    "path": "10.csv",
    "Evento": "Factura electronica 4.0",
    "Mes": "Agosto",
    "Fecha": "04-Agosto-2022",
    "Semana": 20,
    "Hora": "11:30 am",
    "Impartido por": "Ekomercio Electrónico S.A. De C.V."
},
{
    "path": "10.csv",
    "Evento": "Retos que la logística venció durante la pandemia y cómo la logística ayudó a enfrentarlos",
    "Mes": "Agosto",
    "Fecha": "10-Agosto-2022",
    "Semana": 21,
    "Hora": "10:00 am",
    "Impartido por": "Hasar"
},
{
    "path": "10.csv",
    "Evento": "¿Cómo formar embajadores de tus marcas?",
    "Mes": "Agosto",
    "Fecha": "17-Agosto-2022",
    "Semana": 22,
    "Hora": "10:00 am",
    "Impartido por": "Paxia"
},
{
    "path": "10.csv",
    "Evento": "Complemento Carta Porte",
    "Mes": "Agosto",
    "Fecha": "17-Agosto-2022",
    "Semana": 22,
    "Hora": "10:00 am",
    "Impartido por": "Ekomercio Electrónico S.A. De C.V."
},
{
    "path": "10.csv",
    "Evento": "Foro de colaboración Industria Comercio",
    "Mes": "Agosto",
    "Fecha": "18-Agosto-2022",
    "Semana": 22,
    "Hora": "10:00 am",
    "Impartido por": "GS1"
},
{
    "path": "10.csv",
    "Evento": "10 Tendencias de IA Para el nuevo retail",
    "Mes": "Agosto",
    "Fecha": "25-Agosto-2022",
    "Semana": 23,
    "Hora": "10:00 am",
    "Impartido por": "Teamcore"
},
{
    "path": "complemento.csv",
    "Evento": "Complemento Carta Porte",
    "Mes": "Agosto",
    "Fecha": "31-Agosto-2022",
    "Semana": 23,
    "Hora": "11:30 am",
    "Impartido por": "Carvajal"
}]


//TODO: Funcion para obtener los conteos de los archivos

const getInfo = (pathFile, callback) => {
    fs.readFile(pathFile, 'latin1', (error, data) => {
        if(!error) {
            let contador = 0
            const arrayLines = data.split('\n')
            const arrayRegistro = []
            const arrayEmpresas = []
            const arrayAsistentes = []
            const arrayAsis = []
            const arrayEmp = []
            const arrayEmpId = []

            for(const line in arrayLines) {
                let tmp = arrayLines[line].split('|')

                if(contador > 0) {
                    if(!arrayEmpresas.includes(tmp[2]) && tmp[2] != undefined) {
                        arrayEmpresas.push(tmp[2])
                    }

                    if(tmp[4] != undefined) {
                        (!arrayAsistentes.includes(tmp[4])) && (arrayAsistentes.push(tmp[4]));
                        (!arrayRegistro.includes(tmp[4])) && (arrayRegistro.push(tmp[4]));
                    }

                    if(tmp[0] != 'no') {
                        if(!arrayAsis.includes(tmp[4]) && tmp[4] != undefined) {
                            arrayAsis.push(tmp[4])
                        }

                        if(tmp[2] != undefined) {
                            (!arrayEmp.includes(tmp[2])) && (arrayEmp.push(tmp[2]));
                            if(tmp[2] != tmp[10]) {
                                arrayEmpId.push(tmp[10])
                            }
                        }
                    }
                }

                (tmp.length > 1) && (contador ++);
            }
            callback({
                "empresas": arrayEmpresas.length,
                "asistentes": arrayAsistentes.length,
                "registros": arrayRegistro.length,
                "unicos": {
                    "asistentes": arrayAsis.length,
                    "empresas": arrayEmp.length,
                    "id": arrayEmpId
                },
                "total": contador - 1
            })
        }
    })
}

const wb = new xlsx.Workbook()

for(var i in info) {
    getInfo(path.join(__dirname, info[i].path), (rs) => {

        //TODO: Función para modificar el archivo xlsx

        wb.xlsx.readFile(path.join(__dirname, 'demo_eventos.xlsx'))
            .then(() => {
                const ws = wb.getWorksheet(1)
                const lastRow = ws.getRow(ws.actualRowCount)
                console.log(`Ultimo valor: ${lastRow.getCell(1).value}`)
                const row = ws.getRow(ws.actualRowCount + 1)
                row.getCell(1).value = lastRow.getCell(1).value + 1
                row.getCell(2).value = info[i]['Evento']
                row.getCell(3).value = info[i]['Mes']
                row.getCell(4).value = info[i]['Fecha']
                row.getCell(5).value = info[i]['Semana']
                row.getCell(6).value = info[i]['Hora']
                row.getCell(7).value = info[i]['Impartido por']
                row.getCell(8).value = rs.total
                row.getCell(9).value = rs.registros
                row.getCell(10).value = rs.empresas
                row.getCell(11).value = rs.unicos.asistentes
                row.getCell(12).value = rs.unicos.asistentes
                row.getCell(13).value = rs.unicos.empresas
                row.getCell(14).value = rs.unicos.id.length
                row.getCell(15).value = 0
                row.commit()
                wb.xlsx.writeFile(path.join(__dirname, 'demo_eventos.xlsx'))
            })
    })
}
