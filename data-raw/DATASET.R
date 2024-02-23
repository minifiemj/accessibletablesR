## code to prepare `dummydf` dataset goes here

usethis::use_data(DATASET, overwrite = TRUE)

library("tidyverse")

dummydf <- mtcars %>% 
  tibble::rownames_to_column("Car") %>%
  dplyr::rename("Miles per US gallon" = mpg, "Number of cylinders" = cyl, 
                "Displacement\n(cubic inch)" = disp,
                "Gross horsepower" = hp, "Rear axle ratio" = drat, "Weight\n(1000 lbs)" = wt,
                "Quarter mile time" = qsec, "Engine" = vs, "Transmission" = am,
                "Number of forward gears" = gear, "Number of carburettors" = carb) %>%
  dplyr::mutate(Engine = dplyr::case_when(Engine == 0 ~ "V-shaped",
                                          Engine == 1 ~ "Straight")) %>%
  dplyr::mutate(Transmission = dplyr::case_when(Transmission == 0 ~ "Automatic",
                                                Transmission == 1 ~ "Manual")) %>%
  dplyr::mutate("Price\n(£)" = dplyr::case_when(Transmission == "Automatic" ~ "[c]",
                                                Transmission == "Manual" ~ "15767.8752"))