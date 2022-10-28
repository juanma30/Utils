const fs = require('fs')
const path = require('path')
const xlsx = require('xlsx')

const PATH_FILE = path.join(__dirname, '10.csv')

/*
TODO: Función que devuelve la información
      sobre los conteos de los archivos
*/
const createFile = (info, callback) => {
    fs.readFile(PATH_FILE, 'latin1', (error, data) => {
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
                "info": info,
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

const info = {
    "Evento": "10 Tendencias de IA<Plug>PeepOpenara e nuevo retail",
    "Mes": "Agosto",
    "Fecha": "25-Agosto-2022",
    "Semana": 23,
    "Hora": "10:00 am",
    "Impartido por": "Teamcore"
}

createFile(info, (rs) => {
    //TODO: Funcion para crear un xlsx
    console.time('inicio')

    const data = {
        "#": 1,
        "Evento": rs.info['Evento'],
        "Mes": rs.info['Mes'],
        "Fecha": rs.info['Fecha'],
        "Semana": rs.info['Semana'],
        "Hora": rs.info['Hora'],
        "Impartido por": rs.info['Impartido por'],
        "Registros totales": rs.total,
        "Registros únicos": rs.registros,
        "Empresas unicas registradas": rs.empresas,
        "Asistentes Totales": rs.unicos.asistentes,
        "Asistentes únicos": rs.unicos.asistentes,
        "Empresas unicas asistentes": rs.unicos.empresas,
        "Empresas identificadas (Con ID)": rs.unicos.id.length,
        "Empresas sumadas al indicador": 0
    }

    const sheet = xlsx.utils.json_to_sheet([data])

    const archivo = xlsx.utils.book_new()

    xlsx.utils.book_append_sheet(archivo, sheet, 'Reporte 2022')

    xlsx.writeFile(archivo, path.join(__dirname, 'concentrado_eventos_n.xlsx'))

    console.timeEnd('inicio')
})
