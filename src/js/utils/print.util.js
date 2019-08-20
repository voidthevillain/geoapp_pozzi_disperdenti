import { Document, Packer, Paragraph, TextRun, Media, HorizontalPositionAlign, VerticalPositionAlign, HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TextWrappingSide, TextWrappingType } from 'docx'
import { logo_print, scheme1, formula1, scheme2, formula2, formula3, formula4, formula5, scheme3, formula6, formula7 } from './base64.util'

export function createDocument() {
  return (() => {
    const doc = new Document()

    doc.creator = 'Mihai Filip (GeoStru)'
    doc.description = 'Relazione generata dall\'app Pozzi Disperdenti'
    doc.title = 'GA - Pozzi Disperdenti (Relazione)'

    let date = new Date()
    let dd = String(date.getDate()).padStart(2, '0');
    let mm = String(date.getMonth() + 1).padStart(2, '0'); //January is 0!
    let yyyy = date.getFullYear();

    doc.date = '' + dd + '-' + mm + '-' + yyyy

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

    const fig_1 = Media.addImage(doc, scheme1, 304, 268)
    const p_fig_1 = new Paragraph()
    p_fig_1.addImage(fig_1)
    doc.addParagraph(p_fig_1.center())
    
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

    const eq1 = Media.addImage(doc, formula1, 144, 40)
    const p_eq1 = new Paragraph()
    p_eq1.addImage(eq1)
    doc.addParagraph(p_eq1.center())

    const p3 = new Paragraph()
    const p3_1 = new TextRun('Il progetto del pozzo di infiltrazione consiste, essenzialmente, nella determinazione della capacità minima che esso deve avere.').size(24)
    const p3_2 = new TextRun('Questa capacità equivale al volume massimo invasato, che si verifica, come risulta dall’equazione di continuità, quando la portata in smaltimento diventa uguale a quella in entrata.').size(24).break()
    const p3_3 = new TextRun('Riportando in un grafico la portata di piena entrante e quella uscente, in infiltrazione, dal pozzo, il massimo volume d’invaso W').size(24).break()
    const p3_3_1 = new TextRun('0').size(24).subScript()
    const p3_4 = new TextRun(' è dato dall’area compresa tra le due curve, fino al raggiungimento della portata uscente massima Q').size(24)
    const p3_4_1 = new TextRun('f').size(24).subScript()
    const p3_5 = new TextRun(' (Fig. 2).').size(24)
    p3.addRun(p3_1)
    p3.addRun(p3_2)
    p3.addRun(p3_3)
    p3.addRun(p3_3_1)
    p3.addRun(p3_4)
    p3.addRun(p3_4_1)
    p3.addRun(p3_5)
    doc.addParagraph(p3)

    const fig_2 = Media.addImage(doc, scheme2, 304, 233)
    const p_fig_2 = new Paragraph()
    p_fig_2.addImage(fig_2)
    doc.addParagraph(p_fig_2.center())
    
    const fig_caption_2 = new Paragraph()
    const fc2_1 = new TextRun('Fig. 2').bold().size(24)
    const fc2_2 = new TextRun(' - Rappresentazione schematica del processo di laminazione').italics().size(24)
    fig_caption_2.addRun(fc2_1)
    fig_caption_2.addRun(fc2_2)
    doc.addParagraph(fig_caption_2.center())

    const p4 = new Paragraph()
    const p4_1 = new TextRun('I fattori che influiscono sull\'effetto di laminazione operato dalle opere d’invaso sono il volume massimo, in esso contenibile, la sua geometria e la conducibilità idraulica legata alle caratteristiche del terreno.').size(24).break()
    const p4_2 = new TextRun('Il processo di laminazione, nel tempo t è descritto, matematicamente, dal seguente sistema di equazioni:').size(24).break()
    p4.addRun(p4_1)
    p4.addRun(p4_2)
    doc.addParagraph(p4)
 
    const p5 = new Paragraph()
    const p5_1 = new TextRun('  1. Equazione di continuità: ').size(24)
    const eq2 = Media.addImage(doc, formula2, 144, 40)
    const p5_2 = new TextRun('  2. Legge di Darcy: ').size(24).break().break()
    const eq3 = Media.addImage(doc, formula3, 120, 39)
    const p5_3 = new TextRun('  3. Curva d\'invaso: ').size(24).break().break().break()
    const eq4 = Media.addImage(doc, formula4, 98, 21)
    p5.addRun(p5_1)
    p5.addImage(eq2)
    p5.addRun(p5_2)
    p5.addImage(eq3)
    p5.addRun(p5_3)
    p5.addImage(eq4)
    doc.addParagraph(p5)
    
    const p6 = new Paragraph()
    const p6_1 = new TextRun('Una soluzione classica, per pozzi d\'infiltrazione a simmetria assiale, inseriti in un suolo omogeneo, è quella indicata dalla equazione proposta da F. Sieker (Fig. 3): ').size(24).break()
    p6.addRun(p6_1)
    doc.addParagraph(p6)

    const eq5 = Media.addImage(doc, formula5, 256, 80)
    const p_eq5 = new Paragraph()
    p_eq5.addImage(eq5)
    doc.addParagraph(p_eq5.center())

    const p7 = new Paragraph()
    const p7_1 = new TextRun('Dove:').size(24).break()
    const p7_2 = new TextRun('Q').size(24).break()
    const p7_2_1 = new TextRun('f').size(24).subScript()
    const p7_2_2 = new TextRun('   è la portata complessivamente infiltrata [m').size(24)
    const p7_2_3 = new TextRun('3').size(24).superScript()
    const p7_2_4 = new TextRun('/h];').size(24)
    const p7_3 = new TextRun('k/2  è la permeabilità media del terreno insaturo [m/s];').size(24).break()
    const p7_4 = new TextRun('J      è la cadente piezometrica [m/m]; ').size(24).break()
    const p7_5 = new TextRun('L     è la distanza tra la base del pozzo e la superficie di falda [m];').size(24).break()
    const p7_6 = new TextRun('A').size(24).break()
    const p7_6_1 = new TextRun('f').size(24).subScript()
    const p7_6_2 = new TextRun('    è la superficie drenante orizzontale efficace del pozzo, diversa dall’area effettiva della sezione del pozzo A').size(24)
    const p7_6_3 = new TextRun('p').size(24).subScript()
    const p7_6_4 = new TextRun(', di raggio r [m], calcolabile come una corona circolare di larghezza h').size(24)
    const p7_6_5 = new TextRun('w').size(24).subScript()
    const p7_6_6 = new TextRun('/2 dalla quale è escluso l’occludibile fondo [m').size(24)
    const p7_6_7 = new TextRun('2').size(24).superScript()
    const p7_6_8 = new TextRun('];').size(24)
    const p7_7 = new TextRun('h').size(24).break()
    const p7_7_1 = new TextRun('w').size(24).subScript()
    const p7_7_2 = new TextRun('    è il livello idrico nel pozzo [m].').size(24)
    p7.addRun(p7_1)
    p7.addRun(p7_2)
    p7.addRun(p7_2_1)
    p7.addRun(p7_2_2)
    p7.addRun(p7_2_3)
    p7.addRun(p7_2_4)
    p7.addRun(p7_3)
    p7.addRun(p7_4)
    p7.addRun(p7_5)
    p7.addRun(p7_6)
    p7.addRun(p7_6_1)
    p7.addRun(p7_6_2)
    p7.addRun(p7_6_3)
    p7.addRun(p7_6_4)
    p7.addRun(p7_6_5)
    p7.addRun(p7_6_6)
    p7.addRun(p7_6_7)
    p7.addRun(p7_6_8)
    p7.addRun(p7_7)
    p7.addRun(p7_7_1)
    p7.addRun(p7_7_2)
    doc.addParagraph(p7)

    const fig_3 = Media.addImage(doc, scheme3, 304, 445)
    const p_fig_3 = new Paragraph()
    p_fig_3.addImage(fig_3)
    doc.addParagraph(p_fig_3.center())

    const fig_caption_3 = new Paragraph()
    const fc3_1 = new TextRun('Fig. 3').bold().size(24)
    const fc3_2 = new TextRun(' - Schema di pozzo d’infiltrazione secondo F.Sieker').italics().size(24)
    fig_caption_3.addRun(fc3_1)
    fig_caption_3.addRun(fc3_2)
    doc.addParagraph(fig_caption_3.center())

    const p8 = new Paragraph()
    const p8_1 = new TextRun('Il termine ΔW, è espresso, anche, dalla relazione:').size(24).break()
    p8.addRun(p8_1)
    doc.addParagraph(p8)

    const eq6 = Media.addImage(doc, formula6, 88, 45)
    const p_eq6 = new Paragraph()
    p_eq6.addImage(eq6)
    doc.addParagraph(p_eq6)

    const p9 = new Paragraph()
    const p9_1 = new TextRun('dove:').size(24)
    p9.addRun(p9_1)
    doc.addParagraph(p9)

    const eq7 = Media.addImage(doc, formula7, 77, 39)
    const p_eq7 = new Paragraph()
    p_eq7.addImage(eq7)
    doc.addParagraph(p_eq7)

    saveDocument(doc)
  })()
}

function saveDocument(doc) {
  return ((doc) => {
    const packer = new Packer()
    packer.toBlob(doc).then(blob => {
      saveAs(blob, doc.title + ' - ' + doc.date + '.docx')
    })
  })(doc)
}