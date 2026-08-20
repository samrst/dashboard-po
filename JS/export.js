/* ==========================================================
   EXPORT.JS - MÓDULO DE EXPORTAÇÃO PROFISSIONAL v2.3
   Dashboard de Gestão de Avaliações Objetivas SENAI Bahia
   
   Versão: 2.3 - Estabilidade Total e Correção de Dimensionamento
   ========================================================== */

const CONFIG = {
    LOGO_URL: "IMG/senai-logo.png",
    LOGO_WORD_URL: "IMG/senai-logo.png",
    PDF_MARGINS: { top: 15, right: 15, bottom: 15, left: 15 },
    PDF_PAGE_HEIGHT: 297,
    PDF_PAGE_WIDTH: 210,
    PDF_CONTENT_HEIGHT: 270,
    PDF_HEADER_HEIGHT: 30,
    PDF_FOOTER_HEIGHT: 12,
    COLORS: {
        primary: "#003DA5",
        secondary: "#005EB8",
        dark: "#002F80",
        light: "#e8f0fe",
        neutral: "#94a3b8",
        border: "#dce4ef",
        bg_light: "#f5f7fa"
    }
};

const MODAL = {
    overlay: document.getElementById("modal-export"),
    btnOpen: document.getElementById("btn-export"),
    btnClose: document.getElementById("fechar-modal"),
    btnCancel: document.getElementById("cancelar-export"),
    btnConfirm: document.getElementById("confirmar-export"),

    open() {
        if (this.overlay) {
            this.overlay.classList.add("active");
            document.body.style.overflow = "hidden";
        }
    },

    close() {
        if (this.overlay) {
            this.overlay.classList.remove("active");
            document.body.style.overflow = "";
        }
    },

    init() {
        if (this.btnOpen) this.btnOpen.addEventListener("click", () => this.open());
        if (this.btnClose) this.btnClose.addEventListener("click", () => this.close());
        if (this.btnCancel) this.btnCancel.addEventListener("click", () => this.close());
        if (this.overlay) {
            this.overlay.addEventListener("click", (e) => {
                if (e.target === this.overlay) this.close();
            });
        }
    }
};

const UTILS = {
    getElementValue(id) {
        const el = document.getElementById(id);
        return el ? el.textContent.trim() : "0";
    },

    getFilters() {
        const getFilterValue = (id) => {
            const el = document.getElementById(id);
            const value = el?.value;
            if (!value || value === "ALL") return "Todas";
            return value;
        };

        return {
            unidade: getFilterValue("filter-unidade"),
            curso: getFilterValue("filter-curso"),
            modalidade: getFilterValue("filter-modalidade"),
            categoria: getFilterValue("filter-categoria")
        };
    },

    getObservacoes() {
        const el = document.getElementById("export-observacoes");
        return el?.value?.trim() || "";
    },

    getDateTime() {
        return new Date().toLocaleString("pt-BR", {
            day: "2-digit",
            month: "2-digit",
            year: "numeric",
            hour: "2-digit",
            minute: "2-digit"
        });
    },

    getSelectedFormat() {
        const radio = document.querySelector('input[name="tipo"]:checked');
        return radio?.parentElement?.textContent?.trim()?.toUpperCase() || "PDF";
    },

    getSelectedKPIs() {
        const selected = [];
        const allCheckboxes = document.querySelectorAll('.export-option input[type="checkbox"]:checked');
        
        allCheckboxes.forEach(cb => {
            const parent = cb.closest('.export-section');
            if (parent && parent.querySelector('h3')?.textContent.includes('KPIs')) {
                const value = cb.value;
                if (value && value.startsWith("kpi-")) {
                    selected.push(value);
                }
            }
        });
        
        return selected;
    },

    getSelectedCharts() {
        const selected = [];
        const allCheckboxes = document.querySelectorAll('.export-option input[type="checkbox"]:checked');
        
        allCheckboxes.forEach(cb => {
            const value = cb.value;
            if (value && value.startsWith("chart-")) {
                selected.push(value);
            }
        });
        
        return selected;
    },

    loadImage(url) {
        return new Promise((resolve, reject) => {
            const img = new Image();
            img.crossOrigin = "anonymous";
            img.src = url;
            img.onload = () => resolve(img);
            img.onerror = () => reject(new Error(`Erro ao carregar: ${url}`));
        });
    },

    async captureElement(element, scale = 1.5) {
        try {
            if (!window.html2canvas) {
                console.warn("html2canvas não carregado");
                return null;
            }
            
            const canvas = await html2canvas(element, {
                scale: scale,
                backgroundColor: "#ffffff",
                useCORS: true,
                allowTaint: true,
                logging: false
            });
            return canvas.toDataURL("image/png");
        } catch (error) {
            console.error("Erro ao capturar elemento:", error);
            return null;
        }
    }
};

const HEADER = {
    addToPDF(doc, logo, isFirstPage = false) {
        const pageWidth = doc.internal.pageSize.getWidth();
        const margin = CONFIG.PDF_MARGINS.left;
        let y = CONFIG.PDF_MARGINS.top;

        if (logo) {
            doc.addImage(logo, "PNG", margin, y, 18, 10);
        }

        doc.setFont("Helvetica", "bold");
        doc.setFontSize(13);
        doc.setTextColor(0, 61, 165);
        doc.text("Dashboard de Avaliações Objetivas", margin + 22, y + 2);

        doc.setFont("Helvetica", "normal");
        doc.setFontSize(8);
        doc.setTextColor(100, 116, 139);
        doc.text("SENAI Bahia", margin + 22, y + 7);

        y += 12;
        doc.setDrawColor(220, 228, 239);
        doc.line(margin, y, pageWidth - margin, y);

        if (isFirstPage) {
            y += 3;
            doc.setFontSize(7);
            doc.setTextColor(100, 116, 139);
            doc.text(`Gerado em: ${UTILS.getDateTime()}`, margin, y);
        }

        return y + 4;
    },

    addToWord(children, logoBuffer, docxLib) {
        const { Paragraph, HeadingLevel, ImageRun, AlignmentType } = docxLib;

        if (logoBuffer) {
            children.push(new Paragraph({
                children: [
                    new ImageRun({
                        data: logoBuffer,
                        transformation: { width: 80, height: 45 }
                    })
                ],
                alignment: AlignmentType.CENTER
            }));
        }

        children.push(new Paragraph({
            text: "Dashboard de Avaliações Objetivas",
            heading: HeadingLevel.TITLE,
            alignment: AlignmentType.CENTER
        }));

        children.push(new Paragraph({
            text: "SENAI Bahia",
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 }
        }));

        children.push(new Paragraph({
            text: `Gerado em: ${UTILS.getDateTime()}`,
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 }
        }));
    }
};

const FOOTER = {
    addPageNumbers(doc) {
        const pageCount = doc.internal.getNumberOfPages();
        const pageWidth = doc.internal.pageSize.getWidth();
        const margin = CONFIG.PDF_MARGINS.right;
        const footerY = CONFIG.PDF_PAGE_HEIGHT - 10;

        for (let i = 1; i <= pageCount; i++) {
            doc.setPage(i);
            doc.setDrawColor(220, 228, 239);
            doc.line(CONFIG.PDF_MARGINS.left, footerY - 3, pageWidth - margin, footerY - 3);
            doc.setFontSize(7);
            doc.setTextColor(100, 116, 139);
            doc.text("Dashboard SAEP", CONFIG.PDF_MARGINS.left, footerY);
            doc.text("SENAI Bahia", CONFIG.PDF_MARGINS.left + 45, footerY);
            doc.text(`Página ${i} de ${pageCount}`, pageWidth - margin - 15, footerY, { align: "right" });
        }
    },

    addToWord(children, docxLib) {
        const { Paragraph, AlignmentType, BorderStyle } = docxLib;
        children.push(new Paragraph({ text: "" }));
        children.push(new Paragraph({
            text: "Dashboard SAEP — SENAI Bahia",
            alignment: AlignmentType.CENTER,
            spacing: { before: 400 },
            border: {
                top: { color: "dce4ef", space: 1, style: BorderStyle.SINGLE, size: 6 }
            }
        }));
    }
};

const FILTROS = {
    addToPDF(doc, y) {
        const filtros = UTILS.getFilters();
        const margin = CONFIG.PDF_MARGINS.left;
        const pageWidth = doc.internal.pageSize.getWidth();
        const contentWidth = pageWidth - 2 * margin;

        doc.setFont("Helvetica", "bold");
        doc.setFontSize(10);
        doc.setTextColor(0, 61, 165);
        doc.text("Filtros Aplicados", margin, y);
        y += 6;

        doc.setDrawColor(220, 228, 239);
        doc.setFillColor(245, 247, 250);
        const boxHeight = 22;
        doc.rect(margin, y, contentWidth, boxHeight, "FD");

        doc.setFont("Helvetica", "normal");
        doc.setFontSize(8);
        doc.setTextColor(51, 51, 51);

        let filterY = y + 3;
        doc.setFont("Helvetica", "bold");
        doc.setFontSize(7);
        doc.text("Unidade:", margin + 2, filterY);
        doc.setFont("Helvetica", "normal");
        doc.text(filtros.unidade, margin + 25, filterY);

        doc.setFont("Helvetica", "bold");
        doc.text("Curso:", margin + 70, filterY);
        doc.setFont("Helvetica", "normal");
        doc.text(filtros.curso, margin + 85, filterY);

        filterY += 6;
        doc.setFont("Helvetica", "bold");
        doc.text("Modalidade:", margin + 2, filterY);
        doc.setFont("Helvetica", "normal");
        doc.text(filtros.modalidade, margin + 25, filterY);

        doc.setFont("Helvetica", "bold");
        doc.text("Categoria:", margin + 70, filterY);
        doc.setFont("Helvetica", "normal");
        doc.text(filtros.categoria, margin + 85, filterY);

        return y + boxHeight + 8;
    },

    addToWord(children, docxLib) {
        const filtros = UTILS.getFilters();
        const { Paragraph, TextRun } = docxLib;

        children.push(new Paragraph({ text: "Filtros Aplicados", spacing: { before: 200, after: 100 } }));
        children.push(new Paragraph({ children: [new TextRun({ text: `Unidade: `, bold: true }), new TextRun({ text: filtros.unidade })] }));
        children.push(new Paragraph({ children: [new TextRun({ text: `Curso: `, bold: true }), new TextRun({ text: filtros.curso })] }));
        children.push(new Paragraph({ children: [new TextRun({ text: `Modalidade: `, bold: true }), new TextRun({ text: filtros.modalidade })] }));
        children.push(new Paragraph({ children: [new TextRun({ text: `Categoria: `, bold: true }), new TextRun({ text: filtros.categoria })], spacing: { after: 200 } }));
    }
};

const OBSERVACOES = {
    addToPDF(doc, y) {
        const obs = UTILS.getObservacoes();
        if (!obs) return y;

        const margin = CONFIG.PDF_MARGINS.left;
        const pageWidth = doc.internal.pageSize.getWidth();
        const contentWidth = pageWidth - 2 * margin;

        doc.setFont("Helvetica", "bold");
        doc.setFontSize(10);
        doc.setTextColor(0, 61, 165);
        doc.text("Observações", margin, y);
        y += 6;

        doc.setDrawColor(220, 228, 239);
        doc.setFillColor(245, 247, 250);
        const lines = doc.splitTextToSize(obs, contentWidth - 4);
        const boxHeight = Math.min(lines.length * 3.5 + 4, 30);
        doc.rect(margin, y, contentWidth, boxHeight, "FD");
        doc.setFont("Helvetica", "normal");
        doc.setFontSize(7);
        doc.setTextColor(51, 51, 51);
        doc.text(lines, margin + 2, y + 2);

        return y + boxHeight + 8;
    },

    addToWord(children, docxLib) {
        const obs = UTILS.getObservacoes();
        if (!obs) return;
        const { Paragraph, TextRun } = docxLib;
        children.push(new Paragraph({ text: "Observações", spacing: { before: 200, after: 100 } }));
        children.push(new Paragraph({ children: [new TextRun({ text: obs })], spacing: { after: 200 } }));
    }
};

const KPIS = {
    getAll() {
        return [
            { id: "kpi-objetivo", titulo: "Total de Alunos", valor: UTILS.getElementValue("kpi-objetivo"), percentual: "", descricao: "Avaliação Objetiva" },
            { id: "kpi-homologados", titulo: "Alunos Aptos/Agendados", valor: UTILS.getElementValue("kpi-homologados"), percentual: UTILS.getElementValue("kpi-homologados-pct"), descricao: "dos Alunos" },
            { id: "kpi-total", titulo: "Alunos Inaptos/Não Agendados", valor: UTILS.getElementValue("kpi-total"), percentual: UTILS.getElementValue("kpi-total-pct"), descricao: "dos Alunos" },
            { id: "kpi-pendentes", titulo: "Presentes", valor: UTILS.getElementValue("kpi-pendentes"), percentual: UTILS.getElementValue("kpi-pendentes-pct"), descricao: "dos Aptos/Agendados" },
            { id: "kpi-confirmadas", titulo: "Ausentes", valor: UTILS.getElementValue("kpi-confirmadas"), percentual: UTILS.getElementValue("kpi-confirmadas-pct"), descricao: "dos Aptos/Agendados" }
        ];
    },

    getSelected() {
        const selectedIds = UTILS.getSelectedKPIs();
        return this.getAll().filter(kpi => selectedIds.includes(kpi.id));
    },

    addToPDF(doc, y) {
        const kpis = this.getSelected();
        if (kpis.length === 0) return y;

        const margin = CONFIG.PDF_MARGINS.left;
        const pageWidth = doc.internal.pageSize.getWidth();
        const contentWidth = pageWidth - 2 * margin;
        const cardWidth = (contentWidth - 3) / 2;
        const cardHeight = 28;

        doc.setFont("Helvetica", "bold");
        doc.setFontSize(10);
        doc.setTextColor(0, 61, 165);
        doc.text("Resumo Executivo", margin, y);
        y += 6;

        let cardIndex = 0;
        let currentRow = 0;

        while (cardIndex < kpis.length) {
            for (let col = 0; col < 2 && cardIndex < kpis.length; col++) {
                const kpi = kpis[cardIndex];
                const cardX = margin + col * (cardWidth + 3);
                const cardY = y + currentRow * (cardHeight + 2);

                if (cardY + cardHeight > CONFIG.PDF_CONTENT_HEIGHT - 15) {
                    doc.addPage();
                    y = CONFIG.PDF_MARGINS.top + 10;
                    currentRow = 0;
                    col--;
                }

                doc.setDrawColor(220, 228, 239);
                doc.setFillColor(255, 255, 255);
                doc.rect(cardX, cardY, cardWidth, cardHeight, "FD");
                doc.setFillColor(0, 61, 165);
                doc.rect(cardX, cardY, 2, cardHeight, "F");

                doc.setFont("Helvetica", "bold");
                doc.setFontSize(6.5);
                doc.setTextColor(100, 116, 139);
                const titleLines = doc.splitTextToSize(kpi.titulo, cardWidth - 5);
                doc.text(titleLines, cardX + 3, cardY + 2);

                doc.setFont("Helvetica", "bold");
                doc.setFontSize(11);
                doc.setTextColor(0, 61, 165);
                doc.text(String(kpi.valor), cardX + 3, cardY + 10);

                if (kpi.percentual) {
                    doc.setFont("Helvetica", "normal");
                    doc.setFontSize(6);
                    doc.setTextColor(0, 61, 165);
                    doc.text(kpi.percentual, cardX + cardWidth - 3, cardY + 2, { align: "right" });
                }

                doc.setFont("Helvetica", "normal");
                doc.setFontSize(5.5);
                doc.setTextColor(148, 163, 184);
                doc.text(kpi.descricao, cardX + 3, cardY + 16);
                cardIndex++;
            }
            currentRow++;
        }
        return y + currentRow * (cardHeight + 2) + 8;
    },

    addToWord(children, docxLib) {
        const kpis = this.getSelected();
        if (kpis.length === 0) return;
        const { Paragraph, Table, TableCell, TableRow, WidthType, BorderStyle, TextRun } = docxLib;
        children.push(new Paragraph({ text: "Resumo Executivo", spacing: { before: 200, after: 100 } }));
        const rows = [];
        const numRows = Math.ceil(kpis.length / 2);
        for (let i = 0; i < numRows; i++) {
            const cells = [];
            for (let j = 0; j < 2; j++) {
                const kpiIndex = i * 2 + j;
                if (kpiIndex < kpis.length) {
                    const kpi = kpis[kpiIndex];
                    cells.push(new TableCell({
                        children: [
                            new Paragraph({ children: [new TextRun({ text: kpi.titulo, bold: true, size: 18 })], spacing: { after: 80 } }),
                            new Paragraph({ children: [new TextRun({ text: kpi.valor, bold: true, size: 28, color: "003DA5" })], spacing: { after: 60 } }),
                            new Paragraph({ children: [new TextRun({ text: kpi.percentual || "", bold: true, size: 16, color: "003DA5" })], spacing: { after: 60 } }),
                            new Paragraph({ children: [new TextRun({ text: kpi.descricao, size: 14, color: "666666" })], spacing: { after: 0 } })
                        ],
                        borders: { top: { style: BorderStyle.SINGLE, size: 6, color: "dce4ef" }, bottom: { style: BorderStyle.SINGLE, size: 6, color: "dce4ef" }, left: { style: BorderStyle.SINGLE, size: 6, color: "dce4ef" }, right: { style: BorderStyle.SINGLE, size: 6, color: "dce4ef" } },
                        shading: { fill: "f5f7fa" }
                    }));
                } else {
                    cells.push(new TableCell({ children: [new Paragraph({ text: "" })], borders: { top: { style: BorderStyle.SINGLE, size: 6, color: "dce4ef" }, bottom: { style: BorderStyle.SINGLE, size: 6, color: "dce4ef" }, left: { style: BorderStyle.SINGLE, size: 6, color: "dce4ef" }, right: { style: BorderStyle.SINGLE, size: 6, color: "dce4ef" } } }));
                }
            }
            rows.push(new TableRow({ children: cells }));
        }
        children.push(new Table({ rows: rows, width: { size: 100, type: WidthType.PERCENTAGE } }));
        children.push(new Paragraph({ text: "" }));
    }
};

const GRAFICOS = {
    getAll() {
        return [
            { id: "chart-alunos-escola", titulo: "Alunos por Escola" },
            { id: "chart-aplicacao", titulo: "Aptos/Agendados vs Inaptos/Não Agendados" },
            { id: "chart-aplicacao-pizza", titulo: "Visão Geral de Aptidão/Agendamento" },
            { id: "chart-tabulacao", titulo: "Presentes vs Ausentes por Curso" },
            { id: "chart-tabulacao-pizza", titulo: "Eficiência de Presença" },
            { id: "chart-percentual-pratica", titulo: "Participação de Aptos/Agendados" },
            { id: "chart-confirmacao", titulo: "Aptos/Agendados por Curso" }
        ];
    },

    getChartTitle(chartId) {
        const chart = this.getAll().find(c => c.id === chartId);
        return chart ? chart.titulo : "Gráfico";
    },

    async addToPDF(doc, y, chartIds) {
        if (!chartIds || chartIds.length === 0) return y;
        const margin = CONFIG.PDF_MARGINS.left;
        const pageWidth = doc.internal.pageSize.getWidth();
        const contentWidth = pageWidth - 2 * margin;
        let chartNumber = 1;

        doc.setFont("Helvetica", "bold");
        doc.setFontSize(10);
        doc.setTextColor(0, 61, 165);
        doc.text("Gráficos", margin, y);
        y += 6;

        for (const chartId of chartIds) {
            const element = document.getElementById(chartId);
            if (!element) continue;

            if (y + 40 > CONFIG.PDF_CONTENT_HEIGHT - 10) {
                doc.addPage();
                y = CONFIG.PDF_MARGINS.top + 10;
            }

            doc.setFont("Helvetica", "bold");
            doc.setFontSize(9);
            doc.setTextColor(0, 61, 165);
            doc.text(`Gráfico ${String(chartNumber).padStart(2, "0")} - ${this.getChartTitle(chartId)}`, margin, y);
            y += 5;

            const imgData = await UTILS.captureElement(element);
            if (imgData) {
                const img = new Image();
                img.src = imgData;
                await new Promise(r => img.onload = r);
                const aspectRatio = img.width / img.height;
                let imgWidth = contentWidth;
                let imgHeight = contentWidth / aspectRatio;
                if (imgHeight > 50) {
                    imgHeight = 50;
                    imgWidth = 50 * aspectRatio;
                }
                if (y + imgHeight > CONFIG.PDF_CONTENT_HEIGHT - 10) {
                    doc.addPage();
                    y = CONFIG.PDF_MARGINS.top + 10;
                }
                const xOffset = (contentWidth - imgWidth) / 2;
                doc.addImage(imgData, "PNG", margin + xOffset, y, imgWidth, imgHeight);
                y += imgHeight + 3;
                doc.setFont("Helvetica", "normal");
                doc.setFontSize(7);
                doc.setTextColor(100, 116, 139);
                doc.text("Fonte: Dashboard SAEP", margin, y);
                y += 5;
            }
            chartNumber++;
        }
        return y;
    },

    async addToWord(children, chartIds, docxLib) {
        if (!chartIds || chartIds.length === 0) return;
        const { Paragraph, ImageRun, TextRun, AlignmentType } = docxLib;
        children.push(new Paragraph({ text: "Gráficos", spacing: { before: 200, after: 100 } }));
        let chartNumber = 1;
        for (const chartId of chartIds) {
            const element = document.getElementById(chartId);
            if (!element) continue;
            children.push(new Paragraph({ children: [new TextRun({ text: `Gráfico ${String(chartNumber).padStart(2, "0")} - ${this.getChartTitle(chartId)}`, bold: true, size: 20 })], spacing: { before: 200, after: 100 } }));
            const imgData = await UTILS.captureElement(element);
            if (imgData) {
                try {
                    const base64 = imgData.split(",")[1];
                    const buffer = new Uint8Array(atob(base64).split("").map(c => c.charCodeAt(0)));
                    const img = new Image();
                    img.src = imgData;
                    await new Promise(r => img.onload = r);
                    const aspectRatio = img.width / img.height;
                    let imgWidth = 500;
                    let imgHeight = 500 / aspectRatio;
                    if (imgHeight > 350) {
                        imgHeight = 350;
                        imgWidth = 350 * aspectRatio;
                    }
                    children.push(new Paragraph({ children: [new ImageRun({ data: buffer, transformation: { width: Math.round(imgWidth), height: Math.round(imgHeight) } })], alignment: AlignmentType.CENTER, spacing: { after: 100 } }));
                } catch (e) { console.warn("Erro imagem Word:", e); }
            }
            children.push(new Paragraph({ children: [new TextRun({ text: "Fonte: Dashboard SAEP", italics: true, size: 16, color: "666666" })], spacing: { after: 200 } }));
            chartNumber++;
        }
    }
};

const EXPORT_PDF = {
    async generate() {
        try {
            if (!window.jspdf) throw new Error("jsPDF não carregado");
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF("p", "mm", "a4");
            let y = CONFIG.PDF_MARGINS.top;
            let logo = null;
            try { logo = await UTILS.loadImage(CONFIG.LOGO_URL); } catch (e) { console.warn("Logo falhou"); }
            y = HEADER.addToPDF(doc, logo, true);
            y = FILTROS.addToPDF(doc, y);
            y = OBSERVACOES.addToPDF(doc, y);
            y = KPIS.addToPDF(doc, y);
            if (y > CONFIG.PDF_CONTENT_HEIGHT - 50) { doc.addPage(); y = CONFIG.PDF_MARGINS.top + 10; }
            const chartIds = UTILS.getSelectedCharts();
            y = await GRAFICOS.addToPDF(doc, y, chartIds);
            FOOTER.addPageNumbers(doc);
            doc.save("Dashboard-SAEP.pdf");
            return true;
        } catch (error) { alert("Erro PDF: " + error.message); return false; }
    }
};

const EXPORT_WORD = {
    async generate() {
        try {
            if (!window.docx) throw new Error("docx não carregado");
            const docxLib = window.docx;
            const { Document, Packer } = docxLib;
            const children = [];
            let logoBuffer = null;
            try { const r = await fetch(CONFIG.LOGO_WORD_URL); logoBuffer = new Uint8Array(await r.arrayBuffer()); } catch (e) { console.warn("Logo falhou"); }
            HEADER.addToWord(children, logoBuffer, docxLib);
            FILTROS.addToWord(children, docxLib);
            OBSERVACOES.addToWord(children, docxLib);
            KPIS.addToWord(children, docxLib);
            const chartIds = UTILS.getSelectedCharts();
            await GRAFICOS.addToWord(children, chartIds, docxLib);
            FOOTER.addToWord(children, docxLib);
            const doc = new Document({ sections: [{ children }] });
            const blob = await Packer.toBlob(doc);
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url; a.download = "Dashboard-SAEP.docx"; a.click();
            URL.revokeObjectURL(url);
            return true;
        } catch (error) { alert("Erro Word: " + error.message); return false; }
    }
};

const EVENTOS = {
    init() {
        MODAL.init();
        if (MODAL.btnConfirm) {
            MODAL.btnConfirm.addEventListener("click", async () => {
                const formato = UTILS.getSelectedFormat();
                if (!formato) { alert("Selecione um formato."); return; }
                if (UTILS.getSelectedKPIs().length === 0 && UTILS.getSelectedCharts().length === 0) {
                    alert("Selecione algo para exportar."); return;
                }
                MODAL.close();
                this.showProgress();
                try {
                    if (formato === "PDF") await EXPORT_PDF.generate();
                    else if (formato === "WORD") await EXPORT_WORD.generate();
                } catch (e) { console.error(e); } finally { this.hideProgress(); }
            });
        }
        this.setupKPIControls();
        this.setupChartControls();
    },

    setupKPIControls() {
        const btnAll = document.getElementById('btn-select-all-kpis');
        const btnClear = document.getElementById('btn-clear-kpis');
        if (btnAll) btnAll.onclick = () => document.querySelectorAll('.export-section').forEach(s => { if(s.querySelector('h3')?.innerText.includes('KPIs')) s.querySelectorAll('input').forEach(i => i.checked = true); });
        if (btnClear) btnClear.onclick = () => document.querySelectorAll('.export-section').forEach(s => { if(s.querySelector('h3')?.innerText.includes('KPIs')) s.querySelectorAll('input').forEach(i => i.checked = false); });
    },

    setupChartControls() {
        const btnAll = document.getElementById('btn-select-all-charts');
        const btnClear = document.getElementById('btn-clear-charts');
        if (btnAll) btnAll.onclick = () => document.querySelectorAll('.export-section').forEach(s => { if(s.querySelector('h3')?.innerText.includes('Gráficos')) s.querySelectorAll('input').forEach(i => i.checked = true); });
        if (btnClear) btnClear.onclick = () => document.querySelectorAll('.export-section').forEach(s => { if(s.querySelector('h3')?.innerText.includes('Gráficos')) s.querySelectorAll('input').forEach(i => i.checked = false); });
    },

    showProgress() {
        let p = document.getElementById("export-progress");
        if (!p) {
            p = document.createElement("div"); p.id = "export-progress";
            p.innerHTML = `<div class="export-box"><h2>Gerando...</h2><div class="export-progress-bg"><div class="export-progress-fill"></div></div></div>`;
            document.body.appendChild(p);
        }
        p.style.display = "flex";
    },

    hideProgress() {
        const p = document.getElementById("export-progress");
        if (p) p.style.display = "none";
    }
};

document.addEventListener("DOMContentLoaded", () => EVENTOS.init());
