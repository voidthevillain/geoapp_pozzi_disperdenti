import { Document, Packer, Paragraph, TextRun, Media, HorizontalPositionAlign, VerticalPositionAlign, HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TextWrappingSide, TextWrappingType } from 'docx'
import { logo_print, scheme1 } from './base64.util'

export function createDocument() {
  return (() => {
    const doc = new Document()

    doc.creator = 'Mihai Filip (GeoStru)'
    doc.description = 'Relazione generata dall\'app Pozzi Disperdenti'
    doc.title = 'Report'

    const headerImg = Media.addImage(doc, logo_print, 83, 25, {
      floating: {
        horizontalPosition: {
            relative: HorizontalPositionRelativeFrom.PAGE,
            align: HorizontalPositionAlign.CENTER,
        },
        verticalPosition: {
            relative: VerticalPositionRelativeFrom.TOP_MARGIN,
            align: VerticalPositionAlign.CENTER,
        }
    }})
    doc.Header.addImage(headerImg)

    const title = new Paragraph()
    const t_1 = new TextRun('Pozzi Disperdenti').bold()
    title.addRun(t_1)
    doc.addParagraph(title.title().center())
    
    const subtitle = new Paragraph()
    const s_1 = new TextRun('Tecnica dei pozzi superficiali d\'infiltrazione').color('#6D6D6D')
    subtitle.addRun(s_1)
    doc.addParagraph(subtitle.heading2().center())

    const p1 = new Paragraph()
    const p1_1 = new TextRun(`La tecnica dei pozzi superficiali d'infiltrazione (od assorbenti), è adatta al caso di suoli poco permeabili e può essere adoperata per interventi a piccola scala (acque provenienti da tetti isolati) ovvero a media-grande scala (emissari di fognature pluviali).`).size(24).break().break()
    const p1_2 = new TextRun(`Da un punto di vista costruttivo, i pozzi d’infiltrazione sono costituiti da un condotto, senza fondo, che penetra in verticale, sotto la superficie del suolo, in modo da interessare strati particolarmente assorbenti (Fig. 1).`).size(24).break().break()
    p1.addRun(p1_1)
    p1.addRun(p1_2)
    doc.addParagraph(p1)

    const b = new Paragraph()
    const b_1 = new TextRun('').break()
    b.addRun(b_1)
    doc.addParagraph(b)

    const fig_1 = Media.addImage(doc, scheme1, 304, 300, {
      floating: {
        horizontalPosition: {
            relative: HorizontalPositionRelativeFrom.PAGE,
            align: HorizontalPositionAlign.CENTER,
        },
        verticalPosition: {
            relative: VerticalPositionRelativeFrom.LINE,
            align: VerticalPositionAlign.INSIDE,
        },
        wrap: {
          type: TextWrappingType.TOP_AND_BOTTOM,
          side: TextWrappingSide.BOTH_SIDES,
        }}
    })
    doc.addImage(fig_1)
    
    const fig_caption_1 = new Paragraph()
    const fc1_1 = new TextRun('Fig. 1').bold().size(24)
    const fc1_2 = new TextRun(' - Schematizzazione pozzo superficiale d\'infiltrazione').italics().size(24)
    fig_caption_1.addRun(fc1_1)
    fig_caption_1.addRun(fc1_2)
    doc.addParagraph(fig_caption_1.center())

    const p2 = new Paragraph()
    const p2_1 = new TextRun('Da un punto di vista idraulico, i pozzi di infiltrazione sono dei bacini artificiali cilindrici, realizzati allo scopo di smaltire le portate di piena, entro limiti prefissati, dipendenti dalla conducibilità idraulica del terreno.').size(24).break().break()
    const p2_2 = new TextRun('Per operare lo smaltimento e la laminazione delle portate, il pozzo d’infiltrazione deve avere una capacità atta a determinare un processo d\'invaso temporaneo dell\'onda di piena in arrivo ed il suo smaltimento, graduale, nel tempo.').size(24).break().break()
    const p2_3 = new TextRun('Tale processo, di accumulo e laminazione temporale, è descritto, matematicamente, dalla seguente equazione di continuità:').size(24).break().break()
    p2.addRun(p2_1)
    p2.addRun(p2_2)
    p2.addRun(p2_3)
    doc.addParagraph(p2)

    saveDocument(doc)
  })()
}

function saveDocument(doc) {
  return ((doc) => {
    const packer = new Packer()
    packer.toBlob(doc).then(blob => {
      saveAs(blob, doc.title + '.docx')
    })
  })(doc)
}