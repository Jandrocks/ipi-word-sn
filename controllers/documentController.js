const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer, VerticalAlign, AlignmentType, ImageRun } = require('docx');
const fs = require('fs');
const path = require('path');
const { calculateDuration } = require('../utils/helpers');

const generateWordDocument = async (req, res) => {
    const {
        numero,
        empresa,
        cliente,
        correo_cliente,
        telefono_cliente,
        fecha_envio,
        fecha_resuelto,
        ticket_proveedor,
        resumen,
        condicion_falla,
        codigo_resolucion,
        notas_resolucion,
        archivos_adjuntos,
        usuario_asignado,
        observaciones_adicionales,
        notas_trabajo,
        cambios_estado,
        actividades
    } = req.body;

    if (!numero || !empresa || !cliente || !correo_cliente || !telefono_cliente || !fecha_envio || !fecha_resuelto || !resumen || !condicion_falla || !codigo_resolucion || !notas_resolucion) {
        return res.status(400).send('Faltan datos en el cuerpo de la petición.');
    }

    try {
        const duration = calculateDuration(fecha_envio, fecha_resuelto);

        const doc = new Document({
            sections: [
                {
                    headers: {
                        default: new Header({
                            children: [
                                new Paragraph({
                                    children: [
                                        new ImageRun({
                                            data: fs.readFileSync(path.resolve(__dirname, '../public/header_entel.png')),
                                            transformation: {
                                                width: 150,
                                                height: 80,
                                            },
                                        }),
                                    ],
                                    alignment: AlignmentType.RIGHT,
                                }),
                            ],
                        }),
                    },
                    footers: {
                        default: new Footer({
                            children: [
                                new Paragraph({
                                    children: [
                                        new ImageRun({
                                            data: fs.readFileSync(path.resolve(__dirname, '../public/footer.png')),
                                            transformation: {
                                                width: 150,
                                                height: 80,
                                            },
                                        }),
                                    ],
                                    alignment: AlignmentType.CENTER,
                                }),
                            ],
                        }),
                    },
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Información de Cliente",
                                    bold: true,
                                    size: 22,
                                    color: "0000FF",
                                    font: "Arial"
                                }),
                            ],
                        }),
                        new Table({
                            rows: [
                                new TableRow({
                                    children: [
                                        new TableCell({ children: [new Paragraph("Cliente")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph(empresa)], verticalAlign: VerticalAlign.CENTER }),
                                    ],
                                }),
                                new TableRow({
                                    children: [
                                        new TableCell({ children: [new Paragraph("Nombre de contacto cliente")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph(cliente)], verticalAlign: VerticalAlign.CENTER }),
                                    ],
                                }),
                                new TableRow({
                                    children: [
                                        new TableCell({ children: [new Paragraph("Correo electrónico")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph(correo_cliente)], verticalAlign: VerticalAlign.CENTER }),
                                    ],
                                }),
                                new TableRow({
                                    children: [
                                        new TableCell({ children: [new Paragraph("Teléfono")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph(telefono_cliente)], verticalAlign: VerticalAlign.CENTER }),
                                    ],
                                }),
                            ],
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Información de la Incidencia",
                                    bold: true,
                                    size: 22,
                                    color: "0000FF",
                                    font: "Arial"
                                }),
                            ],
                        }),
                        new Table({
                            rows: [
                                new TableRow({
                                    children: [
                                        new TableCell({ children: [new Paragraph("Fecha de envío")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph(fecha_envio)], verticalAlign: VerticalAlign.CENTER }),
                                    ],
                                }),
                                new TableRow({
                                    children: [
                                        new TableCell({ children: [new Paragraph("Fecha de Resuelto")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph(fecha_resuelto)], verticalAlign: VerticalAlign.CENTER }),
                                    ],
                                }),
                                new TableRow({
                                    children: [
                                        new TableCell({ children: [new Paragraph("Duración")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph(duration)], verticalAlign: VerticalAlign.CENTER }),
                                    ],
                                }),
                                new TableRow({
                                    children: [
                                        new TableCell({ children: [new Paragraph("Resumen")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph(resumen)], verticalAlign: VerticalAlign.CENTER }),
                                    ],
                                }),
                                new TableRow({
                                    children: [
                                        new TableCell({ children: [new Paragraph("Condición de Falla")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph(condicion_falla)], verticalAlign: VerticalAlign.CENTER }),
                                    ],
                                }),
                                new TableRow({
                                    children: [
                                        new TableCell({ children: [new Paragraph("Código de Resolución")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph(codigo_resolucion)], verticalAlign: VerticalAlign.CENTER }),
                                    ],
                                }),
                                new TableRow({
                                    children: [
                                        new TableCell({ children: [new Paragraph("Notas de Resolución")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph(notas_resolucion)], verticalAlign: VerticalAlign.CENTER }),
                                    ],
                                }),
                            ],
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Resumen de Actividades Ejecutadas",
                                    bold: true,
                                    size: 22,
                                    color: "0000FF",
                                    font: "Arial"
                                }),
                            ],
                        }),
                        new Table({
                            rows: [
                                new TableRow({
                                    children: [
                                        new TableCell({ children: [new Paragraph("Fecha y Hora")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph("Actividad")], verticalAlign: VerticalAlign.CENTER }),
                                    ],
                                }),
                                // Agregar filas de actividades
                                ...actividades.map(actividad => new TableRow({
                                    children: [
                                        new TableCell({ children: [new Paragraph(actividad.fecha_hora)], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph(actividad.descripcion)], verticalAlign: VerticalAlign.CENTER }),
                                    ],
                                }))
                            ],
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: " ",
                                    bold: true,
                                    size: 22,
                                    color: "0000FF",
                                    font: "Arial"
                                }),
                            ],
                        }),
                        new Table({
                            rows: [
                                new TableRow({
                                    children: [
                                        new TableCell({ children: [new Paragraph("Incidente causado por un cambio")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph("SI / NO")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph("DESCRIPCIÓN DEL CAMBIO")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph("-")], verticalAlign: VerticalAlign.CENTER }),
                                    ],
                                }),
                                new TableRow({
                                    children: [
                                        new TableCell({ children: [new Paragraph("Número de tiquet (proveedor)")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph("")], verticalAlign: VerticalAlign.CENTER }),
                                    ],
                                }),
                            ],
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Diagnóstico Preliminar",
                                    bold: true,
                                    size: 22,
                                    color: "0000FF",
                                    font: "Arial"
                                }),
                            ],
                        }),
                        new Table({
                            rows: [
                                new TableRow({
                                    children: [
                                        new TableCell({
                                            children: [new Paragraph("")],
                                            verticalAlign: VerticalAlign.CENTER,
                                            columnSpan: 2
                                        })
                                    ],
                                })
                            ],
                        })
                    ],
                },
            ],
        });

        // Generar el archivo Word
        const buffer = await Packer.toBuffer(doc);

        // Configurar encabezados para descargar el archivo con la extensión correcta
        res.set({
            'Content-Disposition': `attachment; filename=documento_${Date.now()}.docx`,
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'Content-Length': buffer.length
        });

        // Enviar el archivo al cliente en formato binario
        res.send(buffer);
    } catch (error) {
        console.error('Error al generar el documento:', error);
        res.status(500).send('Error al generar el documento.');
    }
};

module.exports = { generateWordDocument };
