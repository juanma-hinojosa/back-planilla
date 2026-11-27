const express = require('express')
const cors = require('cors');
const ExcelJS = require('exceljs'); // 👈 ¡Faltaba esta línea mágica!

const app = express();
const PORT = 4000;

// Middlewares obligatorios
app.use(cors());
app.use(express.json());

// Datos de prueba (Simulan lo que vendrá de la base de datos o del frontend)
const cotizacionDePrueba = {
  cliente: "Constructora San Marcos",
  proyecto: "Presupuesto Base Estructuras",
  fecha: "2026-06-07",
  rubros: [
    {
      nombreRubro: "Movimiento de Suelos",
      items: [
        { descripcion: "Excavación manual de zanjas", unidad: "m3", cantidad: 15, precioUnitario: 120 },
        { descripcion: "Nivelación y compactación", unidad: "m2", cantidad: 45, precioUnitario: 50 }
      ]
    },
    {
      nombreRubro: "Hormigón Armado",
      items: [
        { descripcion: "Tendido de tuberías de desagüe", unidad: "ml", cantidad: 20, precioUnitario: 80 },
        { descripcion: "Capa aisladora horizontal", unidad: "gl", cantidad: 1, precioUnitario: 400 }
      ]
    }
  ]
};


app.get('/api/exportar-excel', async (req, res) => {
  try {
    // 1. Crear el libro y la hoja de Excel
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Presupuesto');

    // 2. Configurar las columnas (Anchos y claves de acceso)
    worksheet.columns = [
      { header: 'Ítem', key: 'item', width: 10 },
      { header: 'Descripción del Trabajo', key: 'descripcion', width: 45 },
      { header: 'Unidad', key: 'unidad', width: 12 },
      { header: 'Cantidad', key: 'cantidad', width: 15 },
      { header: 'P. Unitario', key: 'precioUnitario', width: 18 },
      { header: 'Total', key: 'total', width: 18 }
    ];

    // 3. Darle estilo a la fila de la cabecera principal (Fila 1)
    const headerRow = worksheet.getRow(1);
    headerRow.height = 25;
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFF' }, size: 11, name: 'Arial' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F4E78' } }; // Azul oscuro
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    let numeroFilaActual = 2; // Empezamos a escribir datos en la fila 2

    // 4. Recorrer los datos dinámicamente para armar los rubros
    cotizacionDePrueba.rubros.forEach((rubro, indexRubro) => {
      const numeroRubro = indexRubro + 1;

      // --- FILA DE RUBRO (Ej: 1.0 Movimiento de Suelos) ---
      const filaRubro = worksheet.addRow([`${numeroRubro}.0`, rubro.nombreRubro.toUpperCase()]);
      filaRubro.height = 22;

      // Combinamos desde la columna B hasta la F para que el título se extienda de forma prolija
      worksheet.mergeCells(`B${numeroFilaActual}:F${numeroFilaActual}`);

      // Estilo visual para destacar la separación de rubros
      filaRubro.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: '000000' }, size: 11, name: 'Arial' };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9E1F2' } }; // Gris azulado claro
        cell.alignment = { vertical: 'middle' };
      });

      numeroFilaActual++;

      // --- FILAS DE ÍTEMS DENTRO DE ESTE RUBRO (Ej: 1.1, 1.2...) ---
      rubro.items.forEach((item, indexItem) => {
        const numeroItem = `${numeroRubro}.${indexItem + 1}`;

        const filaItem = worksheet.addRow({
          item: numeroItem,
          descripcion: item.descripcion,
          unidad: item.unidad,
          cantidad: item.cantidad,
          precioUnitario: item.precioUnitario,
          // Inyectamos la fórmula matemática nativa de Excel multiplicando Cantidad (D) * P.Unitario (E)
          total: { formula: `=D${numeroFilaActual}*E${numeroFilaActual}` }
        });

        filaItem.height = 20;

        // Formatos específicos por celda
        filaItem.getCell('item').alignment = { horizontal: 'center', vertical: 'middle' };
        filaItem.getCell('unidad').alignment = { horizontal: 'center', vertical: 'middle' };
        filaItem.getCell('cantidad').alignment = { horizontal: 'right', vertical: 'middle' };
        
        // Formato de moneda para los precios
        filaItem.getCell('precioUnitario').numFmt = '"$"#,##0.00';
        filaItem.getCell('precioUnitario').alignment = { horizontal: 'right', vertical: 'middle' };
        
        filaItem.getCell('total').numFmt = '"$"#,##0.00';
        filaItem.getCell('total').alignment = { horizontal: 'right', vertical: 'middle' };

        // Añadir bordes inferiores delgados para que parezca una planilla de ingeniería limpia
        filaItem.eachCell((cell) => {
          cell.border = { bottom: { style: 'thin', color: { argb: 'E0E0E0' } } };
          if (!cell.font) cell.font = { name: 'Arial', size: 10 };
        });

        numeroFilaActual++;
      });
    });

    // 5. Configurar las cabeceras HTTP para que el navegador entienda que es un archivo de descarga
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      'attachment; filename=Presupuesto_San_Marcos.xlsx'
    );

    // 6. Escribir el archivo directamente en la respuesta HTTP y cerrar el flujo
    await workbook.xlsx.write(res);
    res.end();

  } catch (error) {
    console.error("Error al generar el Excel:", error);
    res.status(500).json({ mensaje: "Hubo un error interno al generar el archivo" });
  }
});

// Ruta de prueba para verificar que el servidor responda
app.get('/api/saludo', (req, res) => {
  res.json({ mensaje: "¡Backend inicializado con éxito en JavaScript puro!" });
});

// Ruta que devuelve los datos estructurados con tus rubros e ítems
app.get('/api/cotizacion-plantilla', (req, res) => {
  res.json(cotizacionDePrueba);
});

// Iniciar el servidor
app.listen(PORT, () => {
  console.log(`🚀 Servidor JS corriendo en: http://localhost:${PORT}`);
});