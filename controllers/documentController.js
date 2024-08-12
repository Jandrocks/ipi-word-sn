const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer, VerticalAlign, AlignmentType, ImageRun, WidthType } = require('docx');
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
        notas_resolucion,
        archivos_adjuntos,
        usuario_asignado,
        observaciones_adicionales,
        notas_trabajo,
        cambios_estado,
        actividades
    } = req.body;

    const outputFormat = 'base64'; // Cambia a 'base64' para devolver como base64 o a file(para word directo)

    if (!numero || !empresa || !cliente || !correo_cliente || !telefono_cliente || !fecha_envio || !fecha_resuelto || !resumen || !condicion_falla || !notas_resolucion) {
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
                            width: {
                                size: 100,
                                type: WidthType.PERCENTAGE,
                            },
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
                            width: {
                                size: 100,
                                type: WidthType.PERCENTAGE,
                            },
                            rows: [
                                new TableRow({
                                    children: [
                                        new TableCell({ children: [new Paragraph("Fecha del Incidente")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph(fecha_envio)], verticalAlign: VerticalAlign.CENTER }),
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
                                        new TableCell({ children: [new Paragraph("Descripción")], verticalAlign: VerticalAlign.CENTER }),
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
                                        new TableCell({ children: [new Paragraph("Resolución de la Falla")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph(notas_resolucion)], verticalAlign: VerticalAlign.CENTER }),
                                    ],
                                }),
                                new TableRow({
                                    children: [
                                        new TableCell({ children: [new Paragraph("TIPO DOCUMENTO")], verticalAlign: VerticalAlign.CENTER }),
                                        new TableCell({ children: [new Paragraph("CONFIDENCIAL")], verticalAlign: VerticalAlign.CENTER }),
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
                            width: {
                                size: 100,
                                type: WidthType.PERCENTAGE,
                            },
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
                            width: {
                                size: 100,
                                type: WidthType.PERCENTAGE,
                            },
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
                                        new TableCell({ children: [new Paragraph(ticket_proveedor)], verticalAlign: VerticalAlign.CENTER }),
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
                            width: {
                                size: 100,
                                type: WidthType.PERCENTAGE,
                            },
                            rows: [
                                new TableRow({
                                    children: [
                                        new TableCell({
                                            children: [new Paragraph("")],
                                            verticalAlign: VerticalAlign.CENTER,
                                            columnSpan: 2
                                        }),
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

        if (outputFormat === 'file') {
            // Enviar el archivo Word como respuesta para ser descargado
            res.setHeader('Content-Disposition', 'attachment; filename=documento.docx');
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            res.send(buffer);
        } else {
            // Convertir el buffer a base64 y enviarlo
            const base64 = buffer.toString('base64');
            res.json({ base64 });
        }

    } catch (error) {
        console.error('Error al generar el documento:', error); // Registro del error detallado
        res.status(500).send('Error al generar el documento.');
    }
};

module.exports = { generateWordDocument };
