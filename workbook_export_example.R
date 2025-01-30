
# Load packcages ----------------------------------------------------------
library(ggplot2)
library(openxlsx2)

# Load parameters ---------------------------------------------------------


# Load data ---------------------------------------------------------------

penguins <- palmerpenguins::penguins |>
  dplyr::filter(!is.na(sex))


# Fonts config ------------------------------------------------------------

# obter o caminho da font no modo sólido
font_path <- systemfonts::system_fonts() |> 
  dplyr::filter(
    # stringr::str_detect(family, "Roboto"),
    stringr::str_detect(family, "Passion"),
    style == "Regular"
  ) |> 
  dplyr::pull(path)

font_path

# registrar no R para uso no gráfico
systemfonts::register_font(
  # name = "Roboto",
  name = "Passion",
  plain = font_path,
)

# variável para usar no gráfico
# font_family <- "Roboto"
font_family <- "Passion"


# Plot graphic ------------------------------------------------------------

plot <- penguins |>
  ggplot2::ggplot(
    mapping = ggplot2::aes(
      x = species,
      y = bill_length_mm,
      fill = sex
    )
  ) +
  ggplot2::geom_point(
    size = 3,
    alpha = 0.75, 
    shape = 21,
    color = 'black', 
    position = ggplot2::position_jitterdodge(seed = 34543)
  ) +
  ggplot2::theme_minimal(
    base_size = 16,
    base_family = font_family
  ) +
  ggplot2::labs(
    x = ggplot2::element_blank(),
    y = "Comprimento do bico (em mm)",
    fill = "Sexo",
    title = "Medições de diferentes espécies de pinguins"
  )

plot


# Export a workbook 1: Direct Plots ---------------------------------------
# aqui ele não consegue entender a família textual que estou usando, ele salva
# o último gráfico exibido na aba "plots", porém com um tipo de texto genérico

# printar o gráfico desejado
print(plot)

# create a workbook
wb <- wb_workbook()
wb |> 
  openxlsx2::wb_add_worksheet(sheet = "Graphic Test", grid_lines = FALSE) |> 
  openxlsx2::wb_add_data(x = head(penguins)) |> 
  openxlsx2::wb_add_plot(
    dims = "A11",
    width = 16, # Defaults to 6 inches.
    height = 9, # Defaults to 4 in. 
    file_type = "png",
    units = "cm",
    dpi = 300
  ) |>
  openxlsx2::wb_save(file = "TestFile1_DirectPlots.xlsx", overwrite = TRUE)


# Export a workbook 2: Temp files -----------------------------------------

# Criar um arquivo temporário
temp_file <- tempfile(fileext = ".png")

# Salvar gráfico no arquivo temporário
ragg::agg_png(temp_file, width = 16, height = 9, units = "cm", res = 300)
print(plot)
dev.off()


# create a workbook
wb <- wb_workbook()
wb |> 
  openxlsx2::wb_add_worksheet(sheet = "Graphic Test", grid_lines = FALSE) |> 
  openxlsx2::wb_add_data(x = head(penguins)) |> 
  openxlsx2::wb_add_image(
    # file = "plot_temp.png",
    file = temp_file,
    dims = "A11"
  ) |> 
  openxlsx2::wb_save(file = "TestFile2_TempFiles.xlsx", overwrite = TRUE)

# Deletar o arquivo temporário manualmente
unlink(temp_file)



# Export a workbook 3: Images of camcorder --------------------------------

camcorder::gg_record(
  dir ="plots_outputs",
  width = 16,
  height = 9,
  units = "cm",
  dpi = 300,
  bg = "white"
)

# printar o plot
print(plot)

# Encontrar o arquivo mais recente dentro da pasta "plots_outputs"
latest_plot <- base::list.files(
  path = here::here("plots_outputs"),  # Diretório do camcorder
  pattern = "\\.png$",                 # Buscar apenas arquivos PNG
  full.names = TRUE                     # Retornar o caminho completo
) |> 
  tibble::tibble(file_path = _) |>      # Transformar em tibble para manipulação
  dplyr::arrange(dplyr::desc(file_path)) |>  # Ordenar do mais recente para o mais antigo
  dplyr::slice_head(n = 1) |>           # Pegar apenas o mais recente
  dplyr::pull(file_path)                # Extrair o caminho do arquivo

# Adicionar ao Excel
wb |> 
  openxlsx2::wb_add_worksheet(sheet = "Graphic Test", grid_lines = FALSE) |> 
  openxlsx2::wb_add_data(x = head(penguins)) |> 
  openxlsx2::wb_add_image(file = latest_plot, dims = "A11") |> 
  openxlsx2::wb_save(file = "TestFile3_CamCorderImages.xlsx", overwrite = TRUE)


