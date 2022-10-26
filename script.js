const fs = require('fs')
const path = require('path')
const reader = require('xlsx')

const PATH_FILE = path.join(__dirname, 'nuevo_issued.csv')

/*
TODO: Función que devuelve la información 
      sobre los conteos de los archivos
*/
const createFile = (info, callback) => {
    fs.readFile(PATH_FILE, 'latin1', (error, data) => {
        if(!error) {
            let contador = 0
            const arrayLines = data.split('\n')
            const arrayEmpresas = []
            const arrayAsistentes = []
            for(const line in arrayLines) {
                let tmp = arrayLines[line].split('|')
                if(contador > 0) {
                    if(!arrayEmpresas.includes(tmp[9]) && tmp[9] != 'None') {
                        arrayEmpresas.push(tmp[9])
                    }

                    if(!arrayAsistentes.includes(tmp[3])) {
                        arrayAsistentes.push(tmp[3])
                    }
                }
                (tmp.length > 1) && (contador ++)
            }
            callback({
                "info": info,
                "empresas": arrayEmpresas,
                "asistentes": arrayAsistentes,
                "total": contador
            })
        }
    })
}

const info = {
    "Evento": "The Issue",
    "Mes": "Agosto",
    "Fecha": "03-Agosto-2022",
    "Semana": 20,
    "Hora": "10:00 am",
    "Impartido por": "GS1"
}

createFile(info, (rs) => {
    //TODO: Funcion para crear un xlsx
    console.time('inicio')

    const data = [{
        "Evento": rs.info['Evento'],
        "Mes": rs.info['Mes'],
        "Fecha": rs.info['Fecha'],
        "Semana": rs.info['Semana'],
        "Hora": rs.info['Hora'],
        "Impartido por": rs.info['Impartido por'],
        "Registros totales": rs.total,
        "Registros únicos": rs.total,
        "Empresas unicas registradas": rs.empresas.length,
        "Asistentes Totales": 0,
        "Asistentes únicos": 0,
        "Empresas unicas asistentes": rs.asistentes.length,
        "Empresas identificadas (Con ID)": rs.empresas.length,
        "Empresas sumadas al indicador": 0
    }]

    const sheet = reader.utils.json_to_sheet(data)

    const archivo = reader.utils.book_new()

    reader.utils.book_append_sheet(archivo, sheet, 'Reporte 2022')

    reader.writeFile(archivo, path.join(__dirname, 'concentrado_eventos.xlsx'))
    console.timeEnd('inicio')
})
