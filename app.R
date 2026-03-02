# ================================
# app.R (TAM KOD) - Y??l + Ay filtresi, modern renkler, T??rk??e UI (escape),
# KPI + Bar + Donut + Trend + 270 limit + PDF Export (sayfal??)
# ================================

# ---- UTF-8 zorla
try(Sys.setlocale("LC_ALL", "tr_TR.UTF-8"), silent = TRUE)
options(encoding = "UTF-8")

library(shiny)
library(bslib)
library(readxl)
library(dplyr)
library(lubridate)
library(ggplot2)
library(scales)
library(gridExtra)
library(grid)

# ---- T??rk??e aylar (escape ile)
AYLAR_TR <- c(
  "Oca",
  "\u015eub",  # ??ub
  "Mar",
  "Nis",
  "May",
  "Haz",
  "Tem",
  "A\u011fu",  # A??u
  "Eyl",
  "Eki",
  "Kas",
  "Ara"
)

# ---- UI metinleri (escape)
TR <- list(
  TUMU   = "T\u00fcm\u00fc",
  YIL    = "Y\u0131l",
  AY     = "Ay",
  DEPART = "Departman",
  LOKASY = "Lokasyon",
  MESAI  = "Mesai T\u00fcr\u00fc",
  YENILE = "Yenile",
  DASH   = "Fazla Mesai Dashboard",
  TOPLAM = "Toplam Fazla Mesai Saati",
  NORMAL = "Normal Fazla Mesai",
  HAFTA  = "Hafta Sonu Fazla Mesai",
  RESMI  = "Resmi Tatil Fazla Mesai",
  DEVAM  = "Toplam Devams\u0131zl\u0131k Saati",
  BAR    = "Fazla Mesai ve Devams\u0131zl\u0131k",
  DONUT  = "Fazla Mesai T\u00fcrleri",
  TREND  = "Fazla Mesai Trend",
  LIMIT  = "270 Saat Limit Kontrol",
  SAAT   = "Saat"
)

# ---- Excel yolu
path <- "D:/FileSrv/VOLO_Public_Files/Ensar Karayel/fazla_mesai_dashboard.xlsx"

# ---- Modern, kontrast renk paleti
pal <- list(
  fm     = "#2563EB",  # mavi
  dev    = "#EF4444",  # k??rm??z??
  normal = "#10B981",  # turkuaz
  hafta  = "#F59E0B",  # turuncu
  resmi  = "#8B5CF6"   # mor
)

# ================================
# VER?? OKUMA
# ================================
oku_listeler <- function(path) {
  listeler <- read_excel(path, sheet = "Listeler")
  
  dep <- listeler$Departman
  lok <- listeler$Lokasyon
  mes <- listeler$MesaiTuru
  
  dep <- dep[!is.na(dep) & dep != ""]
  lok <- lok[!is.na(lok) & lok != ""]
  mes <- mes[!is.na(mes) & mes != ""]
  
  list(
    departmanlar  = c(TR$TUMU, sort(unique(dep))),
    lokasyonlar   = c(TR$TUMU, sort(unique(lok))),
    mesai_turleri = c(TR$TUMU, sort(unique(mes)))
  )
}

oku_veri <- function(path) {
  read_excel(path, sheet = "Veri") %>%
    mutate(
      Tarih = as.Date(Tarih),
      Yil   = as.integer(year(Tarih)),
      AyNo  = month(Tarih),
      Ay    = factor(AYLAR_TR[AyNo], levels = AYLAR_TR)
    )
}

# ================================
# GRAF??K FONKS??YONLARI (UI + PDF ayn??)
# ================================
plot_bar_fm_dev <- function(d) {
  ozet_wide <- d %>%
    group_by(AyNo, Ay) %>%
    summarise(
      FM  = sum(FazlaMesaiSaat, na.rm = TRUE),
      DEV = sum(DevamsizlikSaat, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    arrange(AyNo)
  
  ozet_long <- bind_rows(
    ozet_wide %>% transmute(AyNo, Ay, Tur = "Fazla Mesai", Saat = FM),
    ozet_wide %>% transmute(AyNo, Ay, Tur = "Devams\u0131zl\u0131k", Saat = DEV)
  ) %>%
    mutate(Label = format(round(Saat, 1), big.mark=".", decimal.mark=","))
  
  ggplot(ozet_long, aes(x = Ay, y = Saat, fill = Tur)) +
    geom_col(position = position_dodge(width = 0.8), width = 0.7) +
    geom_text(
      aes(label = Label),
      position = position_dodge(width = 0.8),
      vjust = -0.35,
      size = 4
    ) +
    scale_fill_manual(values = c("Fazla Mesai" = pal$fm, "Devams\u0131zl\u0131k" = pal$dev)) +
    scale_y_continuous(
      labels = label_number(big.mark=".", decimal.mark=","),
      expand = expansion(mult = c(0, 0.18))
    ) +
    theme_minimal(base_size = 13) +
    theme(
      legend.position = "top",
      plot.title = element_text(face = "bold")
    ) +
    labs(title = TR$BAR, y = TR$SAAT, x = NULL, fill = NULL)
}

plot_donut_mesai <- function(d) {
  tur <- d %>%
    group_by(MesaiTuru) %>%
    summarise(Saat = sum(FazlaMesaiSaat, na.rm = TRUE), .groups = "drop") %>%
    filter(Saat > 0) %>%
    mutate(
      pct = Saat / sum(Saat),
      lbl = percent(pct, accuracy = 1)
    )
  
  if (nrow(tur) == 0) {
    return(ggplot() + theme_void() + labs(title = TR$DONUT, subtitle = "Veri yok"))
  }
  
  renkler <- c(
    "Normal"      = pal$normal,
    "Hafta Sonu"  = pal$hafta,
    "Resmi Tatil" = pal$resmi
  )
  
  ggplot(tur, aes(x = 2, y = Saat, fill = MesaiTuru)) +
    geom_col(color = "white", width = 1) +
    coord_polar("y") +
    xlim(0.6, 2.5) +
    scale_fill_manual(values = renkler) +
    theme_void() +
    labs(title = TR$DONUT) +
    geom_text(
      aes(label = lbl),
      position = position_stack(vjust = 0.5),
      size = 4,
      color = "white"
    )
}

plot_trend_mesai <- function(d) {
  ozet <- d %>%
    group_by(AyNo, Ay, MesaiTuru) %>%
    summarise(Saat = sum(FazlaMesaiSaat, na.rm = TRUE), .groups = "drop") %>%
    arrange(AyNo)
  
  renkler <- c(
    "Normal"      = pal$normal,
    "Hafta Sonu"  = pal$hafta,
    "Resmi Tatil" = pal$resmi
  )
  
  ggplot(ozet, aes(x = Ay, y = Saat, color = MesaiTuru, group = MesaiTuru)) +
    geom_line(linewidth = 1.2) +
    geom_point(size = 2) +
    scale_color_manual(values = renkler) +
    theme_minimal(base_size = 13) +
    labs(title = TR$TREND, y = TR$SAAT, x = NULL, color = NULL) +
    scale_y_continuous(labels = label_number(big.mark=".", decimal.mark=","))
}

# ================================
# UI
# ================================
ui <- page_sidebar(
  title = TR$DASH,
  theme = bs_theme(version = 5, bootswatch = "flatly"),
  tags$head(tags$meta(charset = "utf-8")),
  
  sidebar = sidebar(
    actionButton("yenile", TR$YENILE),
    downloadButton("pdf_indir", "PDF \u0130ndir"),
    hr(),
    selectInput("yil", TR$YIL, choices = NULL),
    checkboxGroupInput(
      "ay",
      TR$AY,
      choices  = AYLAR_TR,
      selected = AYLAR_TR,
      inline   = TRUE
    ),
    selectInput("departman", TR$DEPART, choices = NULL),
    selectInput("lokasyon", TR$LOKASY, choices = NULL),
    selectInput("mesai", TR$MESAI, choices = NULL)
  ),
  
  layout_columns(
    value_box(TR$TOPLAM, textOutput("kpi_toplam")),
    value_box(TR$NORMAL, textOutput("kpi_normal")),
    value_box(TR$HAFTA,  textOutput("kpi_hafta")),
    value_box(TR$RESMI,  textOutput("kpi_resmi")),
    value_box(TR$DEVAM,  textOutput("kpi_dev"))
  ),
  
  layout_columns(
    card(card_header(TR$BAR),   plotOutput("bar_plot",   height = 300)),
    card(card_header(TR$DONUT), plotOutput("donut_plot", height = 300)),
    card(card_header(TR$TREND), plotOutput("trend_plot", height = 300)),
    col_widths = c(4, 4, 4)
  ),
  
  card(
    card_header(TR$LIMIT),
    div(style = "max-height: 320px; overflow-y: auto;",
        tableOutput("limit_table"))
  )
)

# ================================
# SERVER
# ================================
server <- function(input, output, session) {
  
  secenekler <- reactiveVal(NULL)
  df <- reactiveVal(NULL)
  
  observeEvent(TRUE, {
    secenekler(oku_listeler(path))
    df(oku_veri(path))
  }, once = TRUE)
  
  observeEvent(input$yenile, {
    secenekler(oku_listeler(path))
    df(oku_veri(path))
  })
  
  observe({
    req(df(), secenekler())
    
    updateSelectInput(
      session, "yil",
      choices  = sort(unique(df()$Yil)),
      selected = max(df()$Yil, na.rm = TRUE)
    )
    
    updateSelectInput(session, "departman",
                      choices = secenekler()$departmanlar,
                      selected = TR$TUMU)
    
    updateSelectInput(session, "lokasyon",
                      choices = secenekler()$lokasyonlar,
                      selected = TR$TUMU)
    
    updateSelectInput(session, "mesai",
                      choices = secenekler()$mesai_turleri,
                      selected = TR$TUMU)
  })
  
  # filtreli veri
  secili <- reactive({
    req(df(), input$yil)
    
    d <- df() %>% filter(Yil == as.integer(input$yil))
    
    if (!is.null(input$departman) && input$departman != TR$TUMU) {
      d <- d %>% filter(Departman == input$departman)
    }
    if (!is.null(input$lokasyon) && input$lokasyon != TR$TUMU) {
      d <- d %>% filter(Lokasyon == input$lokasyon)
    }
    if (!is.null(input$mesai) && input$mesai != TR$TUMU) {
      d <- d %>% filter(MesaiTuru == input$mesai)
    }
    
    if (!is.null(input$ay) && length(input$ay) > 0) {
      secilen_ay_no <- match(input$ay, AYLAR_TR)
      secilen_ay_no <- secilen_ay_no[!is.na(secilen_ay_no)]
      if (length(secilen_ay_no) > 0) {
        d <- d %>% filter(AyNo %in% secilen_ay_no)
      }
    }
    
    d
  })
  
  # KPI
  output$kpi_toplam <- renderText({
    format(round(sum(secili()$FazlaMesaiSaat, na.rm = TRUE), 1),
           big.mark=".", decimal.mark=",")
  })
  output$kpi_normal <- renderText({
    format(round(sum(secili()$FazlaMesaiSaat[secili()$MesaiTuru == "Normal"], na.rm = TRUE), 1),
           big.mark=".", decimal.mark=",")
  })
  output$kpi_hafta <- renderText({
    format(round(sum(secili()$FazlaMesaiSaat[secili()$MesaiTuru == "Hafta Sonu"], na.rm = TRUE), 1),
           big.mark=".", decimal.mark=",")
  })
  output$kpi_resmi <- renderText({
    format(round(sum(secili()$FazlaMesaiSaat[secili()$MesaiTuru == "Resmi Tatil"], na.rm = TRUE), 1),
           big.mark=".", decimal.mark=",")
  })
  output$kpi_dev <- renderText({
    format(round(sum(secili()$DevamsizlikSaat, na.rm = TRUE), 1),
           big.mark=".", decimal.mark=",")
  })
  
  # UI grafikler
  output$bar_plot <- renderPlot({ plot_bar_fm_dev(secili()) })
  output$donut_plot <- renderPlot({ plot_donut_mesai(secili()) })
  output$trend_plot <- renderPlot({ plot_trend_mesai(secili()) })
  
  # 270 LIMIT tablosu (UI)
  output$limit_table <- renderTable({
    secili() %>%
      group_by(PersonelID, PersonelAdSoyad, Departman, Yil) %>%
      summarise(
        YillikToplamFazlaMesai = sum(FazlaMesaiSaat, na.rm = TRUE),
        .groups = "drop"
      ) %>%
      mutate(LimitAsildiMi = ifelse(YillikToplamFazlaMesai > 270, "EVET", "HAYIR")) %>%
      arrange(desc(YillikToplamFazlaMesai))
  })
  
  # ==========================
  # PDF EXPORT (SAYFALI)
  # ==========================
  output$pdf_indir <- downloadHandler(
    filename = function() {
      paste0("Fazla_Mesai_Rapor_", input$yil, "_", format(Sys.Date(), "%Y%m%d"), ".pdf")
    },
    content = function(file) {
      
      d <- secili()
      
      ay_txt <- if (is.null(input$ay) || length(input$ay) == 0) TR$TUMU else paste(input$ay, collapse = ", ")
      
      filtre_txt <- paste0(
        TR$YIL, ": ", input$yil,
        " | ", TR$AY, ": ", ay_txt,
        " | ", TR$DEPART, ": ", input$departman,
        " | ", TR$LOKASY, ": ", input$lokasyon,
        " | ", TR$MESAI, ": ", input$mesai
      )
      
      # KPI tablo
      kpi_df <- data.frame(
        KPI   = c(TR$TOPLAM, TR$NORMAL, TR$HAFTA, TR$RESMI, TR$DEVAM),
        Deger = c(
          format(round(sum(d$FazlaMesaiSaat, na.rm = TRUE), 1), big.mark=".", decimal.mark=","),
          format(round(sum(d$FazlaMesaiSaat[d$MesaiTuru=="Normal"], na.rm=TRUE), 1), big.mark=".", decimal.mark=","),
          format(round(sum(d$FazlaMesaiSaat[d$MesaiTuru=="Hafta Sonu"], na.rm=TRUE), 1), big.mark=".", decimal.mark=","),
          format(round(sum(d$FazlaMesaiSaat[d$MesaiTuru=="Resmi Tatil"], na.rm=TRUE), 1), big.mark=".", decimal.mark=","),
          format(round(sum(d$DevamsizlikSaat, na.rm=TRUE), 1), big.mark=".", decimal.mark=",")
        ),
        stringsAsFactors = FALSE
      )
      
      # Limit tablo (PDF)
      limit_tbl <- d %>%
        group_by(PersonelID, PersonelAdSoyad, Departman, Yil) %>%
        summarise(YillikToplamFM = sum(FazlaMesaiSaat, na.rm = TRUE), .groups = "drop") %>%
        mutate(LimitAsildiMi = ifelse(YillikToplamFM > 270, "EVET", "HAYIR")) %>%
        arrange(desc(YillikToplamFM))
      
      # Plot objeleri
      p_bar   <- plot_bar_fm_dev(d)
      p_donut <- plot_donut_mesai(d)
      p_trend <- plot_trend_mesai(d)
      
      # PDF a??
      pdf(file, width = 11.69, height = 8.27)  # A4 yatay
      
      # 1) Kapak
      grid.newpage()
      pushViewport(viewport(width = 0.95, height = 0.93))
      gridExtra::grid.arrange(
        grid::textGrob(TR$DASH, gp = grid::gpar(fontsize = 18, fontface = "bold")),
        grid::textGrob(filtre_txt, gp = grid::gpar(fontsize = 10)),
        ncol = 1,
        heights = c(0.25, 0.15)
      )
      
      # 2) KPI sayfas??
      grid.newpage()
      pushViewport(viewport(width = 0.95, height = 0.93))
      grid.text("KPI", y = 0.97, gp = gpar(fontsize = 14, fontface = "bold"))
      gridExtra::grid.arrange(gridExtra::tableGrob(kpi_df, rows = NULL))
      
      # 3) Bar
      grid.newpage()
      pushViewport(viewport(width = 0.95, height = 0.93))
      print(p_bar)
      
      # 4) Donut
      grid.newpage()
      pushViewport(viewport(width = 0.95, height = 0.93))
      print(p_donut)
      
      # 5) Trend
      grid.newpage()
      pushViewport(viewport(width = 0.95, height = 0.93))
      print(p_trend)
      
      # 6+) Limit tablosu: sayfalama + bo??luk
      rows_per_page <- 35
      n <- nrow(limit_tbl)
      pages <- if (n == 0) 1 else ceiling(n / rows_per_page)
      
      for (pg in seq_len(pages)) {
        grid.newpage()
        pushViewport(viewport(width = 0.95, height = 0.93))
        
        grid.text(paste0(TR$LIMIT, " (", pg, "/", pages, ")"),
                  y = 0.97,
                  gp = gpar(fontsize = 14, fontface = "bold"))
        
        if (n == 0) {
          grid.text("Veri yok", y = 0.5)
        } else {
          i1 <- (pg - 1) * rows_per_page + 1
          i2 <- min(pg * rows_per_page, n)
          chunk <- limit_tbl[i1:i2, , drop = FALSE]
          gridExtra::grid.arrange(gridExtra::tableGrob(chunk, rows = NULL))
        }
      }
      
      dev.off()
    }
  )
}

shinyApp(ui, server)