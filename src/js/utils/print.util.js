import { Document, Packer, Paragraph, Table, TextRun, Media, HorizontalPositionAlign, VerticalPositionAlign, HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TextWrappingSide, TextWrappingType, WidthType, ThematicBreak } from 'docx'
import { logo_print, scheme1, formula1, scheme2, formula2, formula3, formula4, formula5, scheme3, formula6, formula7 } from './base64.util'
import { input_k, input_d, input_np, input_as, input_phi, input_tc, input_dp, input_dt, input_a, input_n, table, hwmax, output_hwmax } from './dom.util';

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

    if (output_hwmax.value === '') {
      const p10 = new Paragraph()
      const p10_1 = new TextRun('Non è stato effettuato il calcolo').bold().size(24).break()
      p10.addRun(p10_1)
      doc.addParagraph(p10.center())
    } else {
      const p10 = new Paragraph()
      const p10_1 = new TextRun('Dati').bold().size(24).break().break()
      p10.addRun(p10_1)
      doc.addParagraph(p10)

      const p11 = new Paragraph()
      p11.rightTabStop(5536)
      p11.leftTabStop(5736)
      p11.rightTabStop(6636)
      p11.leftTabStop(6836)
      p11.rightTabStop(7736)
      p11.leftTabStop(7936)
      const p11_1 = new TextRun('Permeabilità media').size(24)
      const p11_2 = new TextRun('K').size(24).tab().tab().tab()
      const p11_3 = new TextRun(input_k.value).bold().size(24).tab().tab()
      const p11_4 = new TextRun('m/s').size(24).tab()
      p11.addRun(p11_1)
      p11.addRun(p11_2)
      p11.addRun(p11_3)
      p11.addRun(p11_4)
      doc.addParagraph(p11)
      
      const p12 = new Paragraph()
      p12.rightTabStop(5536)
      p12.leftTabStop(5736)
      p12.rightTabStop(6636)
      p12.leftTabStop(6836)
      p12.rightTabStop(7736)
      p12.leftTabStop(7936)
      const p12_1 = new TextRun('Diametro pozzetto').size(24)
      const p12_2 = new TextRun('D').size(24).tab().tab().tab()
      const p12_3 = new TextRun(input_d.value).bold().size(24).tab().tab()
      const p12_4 = new TextRun('m').size(24).tab()
      p12.addRun(p12_1)
      p12.addRun(p12_2)
      p12.addRun(p12_3)
      p12.addRun(p12_4)
      doc.addParagraph(p12)

      const p13 = new Paragraph()
      p13.rightTabStop(5536)
      p13.leftTabStop(5736)
      p13.rightTabStop(6636)
      p13.leftTabStop(6836)
      p13.rightTabStop(7736)
      p13.leftTabStop(7936)
      const p13_1 = new TextRun('Numero pozzetti').size(24)
      const p13_2 = new TextRun('N').size(24).tab().tab().tab()
      const p13_3 = new TextRun(input_np.value).bold().size(24).tab().tab()
      const p13_4 = new TextRun('').size(24).tab()
      p13.addRun(p13_1)
      p13.addRun(p13_2)
      p13.addRun(p13_3)
      p13.addRun(p13_4)
      doc.addParagraph(p13)

      const p14 = new Paragraph()
      p14.rightTabStop(5536)
      p14.leftTabStop(5736)
      p14.rightTabStop(6636)
      p14.leftTabStop(6836)
      p14.rightTabStop(7736)
      p14.leftTabStop(7936)
      const p14_1 = new TextRun('Area bacino').size(24)
      const p14_2 = new TextRun('A').size(24).tab().tab().tab()
      const p14_2_1 = new TextRun('s').size(24).subScript()
      const p14_3 = new TextRun(input_as.value).bold().size(24).tab().tab()
      const p14_4 = new TextRun('m').size(24).tab()
      const p14_4_1 = new TextRun('2').size(24).superScript()
      p14.addRun(p14_1)
      p14.addRun(p14_2)
      p14.addRun(p14_2_1)
      p14.addRun(p14_3)
      p14.addRun(p14_4)
      p14.addRun(p14_4_1)
      doc.addParagraph(p14)

      const p15 = new Paragraph()
      p15.rightTabStop(5536)
      p15.leftTabStop(5736)
      p15.rightTabStop(6636)
      p15.leftTabStop(6836)
      p15.rightTabStop(7736)
      p15.leftTabStop(7936)
      const p15_1 = new TextRun('Coefficiente di deflusso').size(24)
      const p15_2 = new TextRun('φ').size(24).tab().tab().tab()
      const p15_3 = new TextRun(input_phi.value).bold().size(24).tab().tab()
      const p15_4 = new TextRun('°').size(24).tab()
      p15.addRun(p15_1)
      p15.addRun(p15_2)
      p15.addRun(p15_3)
      p15.addRun(p15_4)
      doc.addParagraph(p15)

      const p16 = new Paragraph()
      p16.rightTabStop(5536)
      p16.leftTabStop(5736)
      p16.rightTabStop(6636)
      p16.leftTabStop(6836)
      p16.rightTabStop(7736)
      p16.leftTabStop(7936)
      const p16_1 = new TextRun('Tempo di corrivazione').size(24)
      const p16_2 = new TextRun('t').size(24).tab().tab().tab()
      const p16_2_1 = new TextRun('c').size(24).subScript()
      const p16_3 = new TextRun(input_tc.value).bold().size(24).tab().tab()
      const p16_4 = new TextRun('h').size(24).tab()
      p16.addRun(p16_1)
      p16.addRun(p16_2)
      p16.addRun(p16_2_1)
      p16.addRun(p16_3)
      p16.addRun(p16_4)
      doc.addParagraph(p16)

      const p17 = new Paragraph()
      p17.rightTabStop(5536)
      p17.leftTabStop(5736)
      p17.rightTabStop(6636)
      p17.leftTabStop(6836)
      p17.rightTabStop(7736)
      p17.leftTabStop(7936)
      const p17_1 = new TextRun('Durata pioggia').size(24)
      const p17_2 = new TextRun('Δp').size(24).tab().tab().tab()
      const p17_3 = new TextRun(input_dp.value).bold().size(24).tab().tab()
      const p17_4 = new TextRun('h').size(24).tab()
      p17.addRun(p17_1)
      p17.addRun(p17_2)
      p17.addRun(p17_3)
      p17.addRun(p17_4)
      doc.addParagraph(p17)

      const p18 = new Paragraph()
      p18.rightTabStop(5536)
      p18.leftTabStop(5736)
      p18.rightTabStop(6636)
      p18.leftTabStop(6836)
      p18.rightTabStop(7736)
      p18.leftTabStop(7936)
      const p18_1 = new TextRun('Passo integrazione').size(24)
      const p18_2 = new TextRun('Δt').size(24).tab().tab().tab()
      const p18_3 = new TextRun(input_dt.value).bold().size(24).tab().tab()
      const p18_4 = new TextRun('h').size(24).tab()
      p18.addRun(p18_1)
      p18.addRun(p18_2)
      p18.addRun(p18_3)
      p18.addRun(p18_4)
      doc.addParagraph(p18)

      const p19 = new Paragraph()
      p19.rightTabStop(5536)
      p19.leftTabStop(5736)
      p19.rightTabStop(6636)
      p19.leftTabStop(6836)
      p19.rightTabStop(7736)
      p19.leftTabStop(7936)
      const p19_1 = new TextRun('Legge di pioggia h = axt').size(24)
      const p19_1_1 = new TextRun('n').size(24).superScript()
      const p19_2 = new TextRun('a').size(24).tab().tab().tab()
      const p19_3 = new TextRun(input_a.value).bold().size(24).tab().tab()
      const p19_4 = new TextRun('mm/h').size(24).tab()
      p19.addRun(p19_1)
      p19.addRun(p19_1_1)
      p19.addRun(p19_2)
      p19.addRun(p19_3)
      p19.addRun(p19_4)
      doc.addParagraph(p19)

      const p20 = new Paragraph()
      p20.rightTabStop(5536)
      p20.leftTabStop(5736)
      p20.rightTabStop(6636)
      p20.leftTabStop(6836)
      p20.rightTabStop(7736)
      p20.leftTabStop(7936)
      const p20_1 = new TextRun('Legge di pioggia h = axt').size(24)
      const p20_1_1 = new TextRun('n').size(24).superScript()
      const p20_2 = new TextRun('n').size(24).tab().tab().tab()
      const p20_3 = new TextRun(input_n.value).bold().size(24).tab().tab()
      const p20_4 = new TextRun('mm/h').size(24).tab()
      p20.addRun(p20_1)
      p20.addRun(p20_1_1)
      p20.addRun(p20_2)
      p20.addRun(p20_3)
      p20.addRun(p20_4)
      doc.addParagraph(p20)

      const p_break = new Paragraph()
      const p_run = new TextRun('').size(24).break()
      p_break.addRun(p_run)
      doc.addParagraph(p_break)

      const p21 = new Paragraph()
      const p21_1 = new TextRun('Risultati').bold().size(24).break().break()
      p21.addRun(p21_1)
      doc.addParagraph(p21)

      const p22 = new Paragraph()
      p22.rightTabStop(5536)
      p22.leftTabStop(5736)
      p22.rightTabStop(6636)
      p22.leftTabStop(6836)
      p22.rightTabStop(7736)
      p22.leftTabStop(7936)
      const p22_1 = new TextRun('Altezza di progetto del pozzetto').size(24)
      const p22_2 = new TextRun('h').size(24).tab().tab().tab()
      const p22_2_1 = new TextRun('w').size(24).subScript()
      const p22_3 = new TextRun(output_hwmax.value).bold().size(24).tab().tab()
      const p22_4 = new TextRun('m').size(24).tab()
      p22.addRun(p22_1)
      p22.addRun(p22_2)
      p22.addRun(p22_2_1)
      p22.addRun(p22_3)
      p22.addRun(p22_4)
      doc.addParagraph(p22)

      const p_break1 = new Paragraph()
      const p_run1 = new TextRun('').size(24).break()
      p_break1.addRun(p_run1)
      doc.addParagraph(p_break1)

      const pTable = new Table({
        rows: table.rows.length,
        columns: 5,
        width: 100,
        widthUnitType: WidthType.PERCENTAGE
      })
      const pt1 = new Paragraph().center()
      const pt1_1 = new TextRun('t').bold().size(26)
      const pt1_2 = new TextRun('[h]').bold().size(24).break()
      pt1.addRun(pt1_1)
      pt1.addRun(pt1_2)

      const pt2 = new Paragraph().center()
      const pt2_1 = new TextRun('Q').bold().size(26)
      const pt2_1_1 = new TextRun('p').bold().size(26).subScript()
      const pt2_2 = new TextRun('[m').bold().size(24).break()
      const pt2_2_1 = new TextRun('3').bold().size(24).superScript()
      const pt2_2_2 = new TextRun('/h]').bold().size(24)
      pt2.addRun(pt2_1)
      pt2.addRun(pt2_1_1)
      pt2.addRun(pt2_2)
      pt2.addRun(pt2_2_1)
      pt2.addRun(pt2_2_2)

      const pt3 = new Paragraph().center()
      const pt3_1 = new TextRun('Q').bold().size(26)
      const pt3_1_1 = new TextRun('f').bold().size(26).subScript()
      const pt3_2 = new TextRun('[m').bold().size(24).break()
      const pt3_2_1 = new TextRun('3').bold().size(24).superScript()
      const pt3_2_2 = new TextRun('/h]').bold().size(24)
      pt3.addRun(pt3_1)
      pt3.addRun(pt3_1_1)
      pt3.addRun(pt3_2)
      pt3.addRun(pt3_2_1)
      pt3.addRun(pt3_2_2)

      const pt4 = new Paragraph().center()
      const pt4_1 = new TextRun('ΔW').bold().size(26)
      const pt4_2 = new TextRun('[m').bold().size(24).break()
      const pt4_2_1 = new TextRun('3').bold().size(24).superScript()
      const pt4_2_2 = new TextRun(']').bold().size(24)
      pt4.addRun(pt4_1)
      pt4.addRun(pt4_2)
      pt4.addRun(pt4_2_1)
      pt4.addRun(pt4_2_2)

      const pt5 = new Paragraph().center()
      const pt5_1 = new TextRun('h').bold().size(26)
      const pt5_1_1 = new TextRun('w').bold().size(26).subScript()
      const pt5_2 = new TextRun('[m]').bold().size(24).break()
      pt5.addRun(pt5_1)
      pt5.addRun(pt5_1_1)
      pt5.addRun(pt5_2)

      pTable.getRow(0).getCell(0).addParagraph(pt1.center())
      pTable.getRow(0).getCell(1).addParagraph(pt2.center())
      pTable.getRow(0).getCell(2).addParagraph(pt3.center())
      pTable.getRow(0).getCell(3).addParagraph(pt4.center())
      pTable.getRow(0).getCell(4).addParagraph(pt5.center())

      for (let i = 1; i < table.rows.length; i++) {
        let pt6 = new Paragraph().center()
        let pt6_1 = new TextRun('' + table.rows[i].cells[0].innerText).size(24)
        pt6.addRun(pt6_1)
        
        let pt7 = new Paragraph().center()
        let pt7_1 = new TextRun('' + table.rows[i].cells[1].innerText).size(24)
        pt7.addRun(pt7_1)

        let pt8 = new Paragraph().center()
        let pt8_1 = new TextRun('' + table.rows[i].cells[2].innerText).size(24)
        pt8.addRun(pt8_1)

        let pt9 = new Paragraph().center()
        let pt9_1 = new TextRun('' + table.rows[i].cells[3].innerText).size(24)
        pt9.addRun(pt9_1)

        let pt10 = new Paragraph().center()
        let pt10_1 = new TextRun('' + table.rows[i].cells[4].innerText).size(24)
        pt10.addRun(pt10_1)
        
        pTable.getRow(i).getCell(0).addParagraph(pt6.center())
        pTable.getRow(i).getCell(1).addParagraph(pt7.center())
        pTable.getRow(i).getCell(2).addParagraph(pt8.center())
        pTable.getRow(i).getCell(3).addParagraph(pt9.center())
        pTable.getRow(i).getCell(4).addParagraph(pt10.center())
      }

      doc.addTable(pTable)
    }

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