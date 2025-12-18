const XLSX = require('xlsx');
const fs = require('fs');

// Leer archivo de deportes
const deportesWorkbook = XLSX.readFile('deporte.xlsx');
const deportesHojas = ['209', '265', '208'];
const deportesData = {};

deportesHojas.forEach(hoja => {
    if (deportesWorkbook.SheetNames.includes(hoja)) {
        const worksheet = deportesWorkbook.Sheets[hoja];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        deportesData[hoja] = jsonData.map(row => {
            const programa = row['PROGRAMA'] || row['Programa'] || row['programa'] || '';
            const metaProgramada = parseInt(row['META PROGRAMADA'] || row['Meta Programada'] || row['meta programada'] || 0);
            const totalAcumulado = parseInt(row['TOTAL ACUMULADO AÑO'] || row['Total Acumulado Año'] || row['total acumulado año'] || 0);
            let metaPorcentaje = parseFloat(row['META %'] || row['Meta %'] || row['meta %'] || 0);

            if (metaPorcentaje < 1 && metaPorcentaje > 0) {
                metaPorcentaje = metaPorcentaje * 100;
            }

            return {
                programa: programa,
                metaProgramada: metaProgramada,
                totalAcumulado: totalAcumulado,
                metaPorcentaje: parseFloat(metaPorcentaje.toFixed(1))
            };
        });
    }
});

// Leer archivo de infraestructura
const infraWorkbook = XLSX.readFile('infra.xlsx');
const infraHojas = ['207', '220', '201', '214', '217', '007'];
const infraData = {};

infraHojas.forEach(hoja => {
    if (infraWorkbook.SheetNames.includes(hoja)) {
        const worksheet = infraWorkbook.Sheets[hoja];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        infraData[hoja] = jsonData.map(row => {
            const programa = row['PROGRAMA'] || row['Programa'] || row['programa'] || '';
            const metaProgramada = parseInt(row['META PROGRAMADA'] || row['Meta Programada'] || row['meta programada'] || 0);
            const totalMetaAcumulada = parseInt(row['TOTAL META ACUMULADA'] || row['Total Meta Acumulada'] || row['total meta acumulada'] || 0);
            let metaPorcentaje = parseFloat(row['META %'] || row['Meta %'] || row['meta %'] || 0);
            const link = row['LINK DE ACCESO DRIVE( EVIDENCIAS)'] || row['Link de acceso drive'] || row['link'] || '#';

            if (metaPorcentaje < 1 && metaPorcentaje > 0) {
                metaPorcentaje = metaPorcentaje * 100;
            }

            return {
                programa: programa,
                metaProgramada: metaProgramada,
                totalMetaAcumulada: totalMetaAcumulada,
                metaPorcentaje: parseFloat(metaPorcentaje.toFixed(1)),
                link: link
            };
        });
    }
});

// Crear el objeto de datos completo
const data = {
    deportesData: deportesData,
    infraData: infraData
};

// Guardar como JSON
fs.writeFileSync('data.json', JSON.stringify(data, null, 2));

console.log('✓ Archivos Excel convertidos a JSON exitosamente');
console.log('- Deportes:', Object.keys(deportesData).length, 'hojas');
console.log('- Infraestructura:', Object.keys(infraData).length, 'hojas');
