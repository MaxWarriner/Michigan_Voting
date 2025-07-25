library(dplyr)
library(tidyverse)
library(openxlsx)  # For Excel writing
library(ggplot2)   # For plotting
library(patchwork) # For arranging multiple ggplots
library(stringr)

dat <- read.csv("1998_2024_data.csv")

dat <- dat |>
  mutate(dem = ifelse(Party == "DEM", 'DEM', 'OTHER'), 
         City_Name = str_replace(City_Name, "ST\\.", "ST")) |>
  filter(!grepl("CHARTER TOWNSHIP", City_Name))

  for(county in unique(dat$County_Name)[36:length(unique(dat$County_Name))]){

  ######################################################
  # Executive Office Data
  ######################################################
  pres_gov <- dat |>
    dplyr::filter(County_Name == county, 
                  Office == "President" | Office == "Governor", 
                  City_Name != "{Statistical Adjustments}") |>
    dplyr::select(City_Name, dem, Votes, Year) |>
    group_by(City_Name, dem, Year) |>
    reframe(Votes = sum(Votes)) |>
    pivot_wider(names_from = dem, values_from = Votes) |>
    mutate(dem_percent = round((DEM)/(DEM + OTHER),4),
           turnout = DEM + OTHER) |>
    dplyr::select(-OTHER) |>
    pivot_wider(names_from = Year, values_from = c(DEM, dem_percent, turnout))
  
  pres_gov_total <- dat |>
    dplyr::filter(County_Name == county, 
                                        Office == "President" | Office == "Governor", 
                                        City_Name != "{Statistical Adjustments}") |>
    dplyr::select(City_Name, dem, Votes, Year) |>
    group_by(dem, Year) |>
    reframe(Votes = sum(Votes)) |>
    pivot_wider(names_from = dem, values_from = Votes) |>
    mutate(dem_percent = round((DEM)/(DEM + OTHER),4),
           turnout = DEM + OTHER) |>
    dplyr::select(-OTHER) |>
    pivot_wider(names_from = Year, values_from = c(DEM, dem_percent, turnout)) |>
    mutate(City_Name = "Total")
    
    pres_gov <- bind_rows(pres_gov, pres_gov_total)
  
    pres_gov <- pres_gov[,c(1,2, 16, 30, 3, 17, 31, 4, 18, 32, 5, 19, 33, 6, 20, 34, 7, 21, 35, 8, 22, 36, 9, 23, 37, 10, 24, 38, 11, 25, 39, 12, 26, 40, 13, 27, 41, 14, 28, 42, 15, 29, 43)]
  
  
  
  
  ######################################################
  # Executive Office Plots
  ######################################################
  
  executive_plots = list()
  
  for(i in 1:(length(unique(pres_gov$City_Name))-1)){
    
    plot_data <- dat |>
      dplyr::filter(County_Name == county, 
                    Office == "President" | Office == "Governor", 
                    City_Name != "{Statistical Adjustments}",
                    City_Name == pres_gov$City_Name[i]) |>
      dplyr::select(City_Name, dem, Votes, Year) |>
      group_by(City_Name, dem, Year) |>
      reframe(Votes = sum(Votes)) |>
      pivot_wider(names_from = dem, values_from = Votes)
    if(ncol(plot_data) == 4){
    plot_data <- plot_data |>
      mutate(dem_percent = round((DEM)/(DEM + OTHER),2),
             turnout = DEM + OTHER) |>
      dplyr::select(-OTHER)
    }else{
      plot_data <- plot_data |>
        mutate(DEM = rep(0, nrow(plot_data))) |>
        mutate(dem_percent = round((DEM)/(DEM + OTHER),2),
               turnout = DEM + OTHER) |>
        dplyr::select(-OTHER)
    }
    
    executive_plots[[i]] <- ggplot(data = plot_data, aes(x = Year, y = dem_percent)) + 
      geom_line() +
      geom_point() + 
      geom_smooth(method = lm, alpha = 0.2, se = FALSE) + 
      geom_text(aes(label = dem_percent), nudge_y = 0.08, size = 2) + 
      theme_bw() +
      geom_hline(yintercept = 0) +
      scale_x_continuous(breaks = unique(plot_data$Year)) + 
      ggtitle(paste(pres_gov$City_Name[i], ": President/Governor", sep = "")) +
      ylim(0,1) + 
      theme(plot.title = element_text(hjust = 0.5, size = 8), 
            axis.title.x = element_text(size = 8), 
            axis.title.y = element_text(size = 8), 
            axis.text.x = element_text(size = 6), 
            axis.text.y = element_text(size = 6)) +
      ylab("Dem Percent")
    
  }
  
  executive_plot_total <- dat |>
    dplyr::filter(County_Name == county, 
                  Office == "President" | Office == "Governor", 
                  City_Name != "{Statistical Adjustments}") |>
    dplyr::select(City_Name, dem, Votes, Year) |>
    group_by(dem, Year) |>
    reframe(Votes = sum(Votes)) |>
    pivot_wider(names_from = dem, values_from = Votes) |>
    mutate(dem_percent = round((DEM)/(DEM + OTHER),2),
           turnout = DEM + OTHER) |>
    dplyr::select(-OTHER) |>
    mutate(City_Name = "Total")
  
  
  executive_plots[[length(unique(pres_gov$City_Name))]] <- ggplot(data = executive_plot_total, aes(x = Year, y = dem_percent)) + 
    geom_line() +
    geom_point() + 
    geom_smooth(method = lm, alpha = 0.2, se = FALSE) + 
    geom_text(aes(label = dem_percent), nudge_y = 0.08, size = 2) + 
    theme_bw() +
    geom_hline(yintercept = 0) +
    scale_x_continuous(breaks = unique(plot_data$Year)) + 
    ggtitle("Total: President & Governor") +
    ylim(0,1) + 
    theme(plot.title = element_text(hjust = 0.5, size = 8), 
          axis.title.x = element_text(size = 8), 
          axis.title.y = element_text(size = 8), 
          axis.text.x = element_text(size = 6), 
          axis.text.y = element_text(size = 6)) +
    ylab("Dem Percent")
  
  executive_plots <- wrap_plots(executive_plots, ncol = 3)

  ######################################################
  # State House Data
  ######################################################
  state_house <- dat |>
    dplyr::filter(County_Name == county, 
                  Office == "State Rep", 
                  City_Name != "{Statistical Adjustments}") |>
    dplyr::select(City_Name, dem, Votes, Year) |>
    group_by(City_Name, dem, Year) |>
    reframe(Votes = sum(Votes)) |>
    pivot_wider(names_from = dem, values_from = Votes) |>
    mutate(dem_percent = round((DEM)/(DEM + OTHER),4),
           turnout = DEM + OTHER) |>
    dplyr::select(-OTHER) |>
    pivot_wider(names_from = Year, values_from = c(DEM, dem_percent, turnout))
  
  state_housetotal <- dat |>
    dplyr::filter(County_Name == county, 
                  Office == "State Rep", 
                  City_Name != "{Statistical Adjustments}") |>
    dplyr::select(City_Name, dem, Votes, Year) |>
    group_by(dem, Year) |>
    reframe(Votes = sum(Votes)) |>
    pivot_wider(names_from = dem, values_from = Votes) |>
    mutate(dem_percent = round((DEM)/(DEM + OTHER),4),
           turnout = DEM + OTHER) |>
    dplyr::select(-OTHER) |>
    pivot_wider(names_from = Year, values_from = c(DEM, dem_percent, turnout)) |>
    mutate(City_Name = "Total")
  
  state_house <- bind_rows(state_house, state_housetotal)
  
  state_house <- state_house[,c(1,2, 16, 30, 3, 17, 31, 4, 18, 32, 5, 19, 33, 6, 20, 34, 7, 21, 35, 8, 22, 36, 9, 23, 37, 10, 24, 38, 11, 25, 39, 12, 26, 40, 13, 27, 41, 14, 28, 42, 15, 29, 43)]
  
  ######################################################
  # State House Plots
  ######################################################
  
  state_house_plots = list()
  
  for(i in 1:(length(unique(state_house$City_Name))-1)){
    
    plot_data <- dat |>
      dplyr::filter(County_Name == county, 
                    Office == "State Rep", 
                    City_Name != "{Statistical Adjustments}",
                    City_Name == state_house$City_Name[i]) |>
      dplyr::select(City_Name, dem, Votes, Year) |>
      group_by(City_Name, dem, Year) |>
      reframe(Votes = sum(Votes)) |>
      pivot_wider(names_from = dem, values_from = Votes)
    
      if(ncol(plot_data) == 4){
        plot_data <- plot_data |>
          mutate(dem_percent = round((DEM)/(DEM + OTHER),2),
                 turnout = DEM + OTHER) |>
          dplyr::select(-OTHER)
      }else{
        plot_data <- plot_data |>
          mutate(DEM = rep(0, nrow(plot_data))) |>
          mutate(dem_percent = round((DEM)/(DEM + OTHER),2),
                 turnout = DEM + OTHER) |>
          dplyr::select(-OTHER)
      }
    
    
    state_house_plots[[i]] <- ggplot(data = plot_data, aes(x = Year, y = dem_percent)) + 
      geom_line() +
      geom_point() + 
      geom_smooth(method = lm, alpha = 0.2, se = FALSE) + 
      geom_text(aes(label = dem_percent), nudge_y = 0.08, size = 2) + 
      theme_bw() +
      geom_hline(yintercept = 0) +
      scale_x_continuous(breaks = unique(plot_data$Year)) + 
      ggtitle(paste(pres_gov$City_Name[i], ": State House", sep = "")) +
      ylim(0,1) + 
      theme(plot.title = element_text(hjust = 0.5, size = 8), 
            axis.title.x = element_text(size = 8), 
            axis.title.y = element_text(size = 8), 
            axis.text.x = element_text(size = 6), 
            axis.text.y = element_text(size = 6)) +
      ylab("Dem Percent")
    
  }
  
  state_house_plot_total <- dat |>
    dplyr::filter(County_Name == county, 
                  Office == "State Rep", 
                  City_Name != "{Statistical Adjustments}") |>
    dplyr::select(City_Name, dem, Votes, Year) |>
    group_by(dem, Year) |>
    reframe(Votes = sum(Votes)) |>
    pivot_wider(names_from = dem, values_from = Votes) |>
    mutate(dem_percent = round((DEM)/(DEM + OTHER),2),
           turnout = DEM + OTHER) |>
    dplyr::select(-OTHER) |>
    mutate(City_Name = "Total")
  
  
  state_house_plots[[length(unique(pres_gov$City_Name))]] <- ggplot(data = state_house_plot_total, aes(x = Year, y = dem_percent)) + 
    geom_line() +
    geom_point() + 
    geom_smooth(method = lm, alpha = 0.2, se = FALSE) + 
    geom_text(aes(label = dem_percent), nudge_y = 0.08, size = 2) + 
    theme_bw() +
    geom_hline(yintercept = 0) +
    scale_x_continuous(breaks = unique(plot_data$Year)) + 
    ggtitle("Total: State House") +
    ylim(0,1) + 
    theme(plot.title = element_text(hjust = 0.5, size = 8), 
          axis.title.x = element_text(size = 8), 
          axis.title.y = element_text(size = 8), 
          axis.text.x = element_text(size = 6), 
          axis.text.y = element_text(size = 6)) +
    ylab("Dem Percent")
  
  
  
  state_house_plots <- wrap_plots(state_house_plots, ncol = 3)
  
  ######################################################
  # National Congress Data
  ######################################################
  congress <- dat |>
    dplyr::filter(County_Name == county, 
                  Office == "U.S. Rep", 
                  City_Name != "{Statistical Adjustments}") |>
    dplyr::select(City_Name, dem, Votes, Year) |>
    group_by(City_Name, dem, Year) |>
    reframe(Votes = sum(Votes)) |>
    pivot_wider(names_from = dem, values_from = Votes) |>
    mutate(dem_percent = round((DEM)/(DEM + OTHER),4),
           turnout = DEM + OTHER) |>
    dplyr::select(-OTHER) |>
    pivot_wider(names_from = Year, values_from = c(DEM, dem_percent, turnout))
  
  congress_total <- dat |>
    dplyr::filter(County_Name == county, 
                  Office == "U.S. Rep", 
                  City_Name != "{Statistical Adjustments}") |>
    dplyr::select(City_Name, dem, Votes, Year) |>
    group_by(dem, Year) |>
    reframe(Votes = sum(Votes)) |>
    pivot_wider(names_from = dem, values_from = Votes) |>
    mutate(dem_percent = round((DEM)/(DEM + OTHER),4),
           turnout = DEM + OTHER) |>
    dplyr::select(-OTHER) |>
    pivot_wider(names_from = Year, values_from = c(DEM, dem_percent, turnout)) |>
    mutate(City_Name = "Total")
  
  congress <- bind_rows(congress, congress_total)
  
  congress <- congress[,c(1,2, 16, 30, 3, 17, 31, 4, 18, 32, 5, 19, 33, 6, 20, 34, 7, 21, 35, 8, 22, 36, 9, 23, 37, 10, 24, 38, 11, 25, 39, 12, 26, 40, 13, 27, 41, 14, 28, 42, 15, 29, 43)]
  
  ######################################################
  # National Congress Plots
  ######################################################
  
  congress_plots = list()
  
  for(i in 1:(length(unique(congress$City_Name))-1)){
    
    plot_data <- dat |>
      dplyr::filter(County_Name == county, 
                    Office == "U.S. Rep", 
                    City_Name != "{Statistical Adjustments}",
                    City_Name == pres_gov$City_Name[i]) |>
      dplyr::select(City_Name, dem, Votes, Year) |>
      group_by(City_Name, dem, Year) |>
      reframe(Votes = sum(Votes)) |>
      pivot_wider(names_from = dem, values_from = Votes)
      if(ncol(plot_data) == 4){
        plot_data <- plot_data |>
          mutate(dem_percent = round((DEM)/(DEM + OTHER),2),
                 turnout = DEM + OTHER) |>
          dplyr::select(-OTHER)
      }else{
        plot_data <- plot_data |>
          mutate(DEM = rep(0, nrow(plot_data))) |>
          mutate(dem_percent = round((DEM)/(DEM + OTHER),2),
                 turnout = DEM + OTHER) |>
          dplyr::select(-OTHER)
      }
    
    
    congress_plots[[i]] <- ggplot(data = plot_data, aes(x = Year, y = dem_percent)) + 
      geom_line() +
      geom_point() + 
      geom_smooth(method = lm, alpha = 0.2, se = FALSE) + 
      geom_text(aes(label = dem_percent), nudge_y = 0.08, size = 2) + 
      theme_bw() +
      geom_hline(yintercept = 0) +
      scale_x_continuous(breaks = unique(plot_data$Year)) + 
      ggtitle(paste(pres_gov$City_Name[i], ": President/Governor", sep = "")) +
      ylim(0,1) + 
      theme(plot.title = element_text(hjust = 0.5, size = 8), 
            axis.title.x = element_text(size = 8), 
            axis.title.y = element_text(size = 8), 
            axis.text.x = element_text(size = 6), 
            axis.text.y = element_text(size = 6)) +
      ylab("Dem Percent")
  }
  
  congress_plot_total <- dat |>
    dplyr::filter(County_Name == county, 
                  Office == "U.S. Rep", 
                  City_Name != "{Statistical Adjustments}") |>
    dplyr::select(City_Name, dem, Votes, Year) |>
    group_by(dem, Year) |>
    reframe(Votes = sum(Votes)) |>
    pivot_wider(names_from = dem, values_from = Votes) |>
    mutate(dem_percent = round((DEM)/(DEM + OTHER),4),
           turnout = DEM + OTHER) |>
    dplyr::select(-OTHER) |>
    mutate(City_Name = "Total")
  
  congress_plots[[length(unique(pres_gov$City_Name))]] <- ggplot(data = congress_plot_total, aes(x = Year, y = dem_percent)) + 
    geom_line() +
    geom_point() + 
    geom_smooth(method = lm, alpha = 0.2, se = FALSE) + 
    geom_text(aes(label = dem_percent), nudge_y = 0.08, size = 2) + 
    theme_bw() +
    geom_hline(yintercept = 0) +
    scale_x_continuous(breaks = unique(plot_data$Year)) + 
    ggtitle("Total: State House") +
    ylim(0,1) + 
    theme(plot.title = element_text(hjust = 0.5, size = 8), 
          axis.title.x = element_text(size = 8), 
          axis.title.y = element_text(size = 8), 
          axis.text.x = element_text(size = 6), 
          axis.text.y = element_text(size = 6)) +
    ylab("Dem Percent")
  
  congress_plots <- wrap_plots(congress_plots, ncol = 3)
  #######################################################
  # Compiled Plots
  #######################################################
  compiled_plots = list()
  
  for(i in 1:(length(unique(congress$City_Name))-1)){
    
    plot_data <- dat |>
      dplyr::filter(County_Name == county, 
                    Office == "U.S. Rep" | Office == "President" | Office == "Governor" | Office == "State Rep", 
                    City_Name != "{Statistical Adjustments}",
                    City_Name == pres_gov$City_Name[i]) |>
      dplyr::select(City_Name, dem, Votes, Year, Office) |>
      group_by(City_Name, dem, Year, Office) |>
      reframe(Votes = sum(Votes)) |>
      pivot_wider(names_from = dem, values_from = Votes)
      if(ncol(plot_data) == 5){
        plot_data <- plot_data |>
          mutate(dem_percent = round((DEM)/(DEM + OTHER),2),
                 turnout = DEM + OTHER) |>
          dplyr::select(-OTHER)
      }else{
        plot_data <- plot_data |>
          mutate(DEM = rep(0, nrow(plot_data))) |>
          mutate(dem_percent = round((DEM)/(DEM + OTHER),2),
                 turnout = DEM + OTHER) |>
          dplyr::select(-OTHER)
      }
    
    compiled_plots[[i]] <- ggplot(data = plot_data, aes(x = Year, y = dem_percent, colour = Office)) + 
      geom_line() +
      geom_point() + 
      theme_bw() +
      geom_hline(yintercept = 0) +
      scale_x_continuous(breaks = unique(plot_data$Year)) + 
      ggtitle(paste(pres_gov$City_Name[i], ": Compiled", sep = "")) +
      ylim(0,1) + 
      theme(plot.title = element_text(hjust = 0.5, size = 8), 
            axis.title.x = element_text(size = 8), 
            axis.title.y = element_text(size = 8), 
            axis.text.x = element_text(size = 6), 
            axis.text.y = element_text(size = 6),
            legend.text = element_text(size = 8), 
            legend.title = element_text(size = 10)) +
      ylab("Dem Percent")
  }
  
  compiled_plot_total <- dat |>
    dplyr::filter(County_Name == county, 
                  Office == "U.S. Rep" | Office == "President" | Office == "Governor" | Office == "State Rep", 
                  City_Name != "{Statistical Adjustments}") |>
    dplyr::select(Office, dem, Votes, Year) |>
    group_by(dem, Year, Office) |>
    reframe(Votes = sum(Votes)) |>
    pivot_wider(names_from = dem, values_from = Votes) |>
        mutate(dem_percent = round((DEM)/(DEM + OTHER),2),
               turnout = DEM + OTHER) |>
        dplyr::select(-OTHER)

  
  compiled_plots[[length(unique(pres_gov$City_Name))]] <- ggplot(data = compiled_plot_total, aes(x = Year, y = dem_percent, colour = Office)) + 
    geom_line() +
    geom_point() + 
    geom_smooth(method = lm, alpha = 0.2, se = FALSE) + 
    theme_bw() +
    geom_hline(yintercept = 0) +
    scale_x_continuous(breaks = unique(plot_data$Year)) + 
    ggtitle("Total: Compiled") +
    ylim(0,1) + 
    theme(plot.title = element_text(hjust = 0.5, size = 8), 
          axis.title.x = element_text(size = 8), 
          axis.title.y = element_text(size = 8), 
          axis.text.x = element_text(size = 6), 
          axis.text.y = element_text(size = 6)) +
    ylab("Dem Percent")
  
  
  
  compiled_plots <- wrap_plots(compiled_plots, ncol = 3) +
    plot_layout(guides = "collect") & theme(legend.position = "top", legend.direction = "horizontal")

  #######################################################
  # Create Spreadsheet
  #######################################################
  wb <- createWorkbook()
  
  addWorksheet(wb, "President and Governor")
  writeData(wb, "President and Governor", pres_gov)
  
  addWorksheet(wb, "US Congress")
  writeData(wb, "US Congress", congress)
  
  addWorksheet(wb, "State House")
  writeData(wb, "State House", state_house)
  
  # Executive Collage
  addWorksheet(wb, "President and Governor Plots")
  
  rows = round(length(executive_plots) / 3)
  
  temp_executive_plot <- tempfile(fileext = ".png")
  ggsave(temp_executive_plot, executive_plots, width = 12, height = rows * 3, units = "in", dpi = 700, limitsize = FALSE)
  
  
  insertImage(
    wb, 
    sheet = "President and Governor Plots", 
    file = temp_executive_plot, 
    width = 12,  # Adjust size as needed
    height = rows*3,
    startRow = 1,  # Position in Excel (row 1, column 1)
    startCol = 1,
    units = "in"   # Dimensions in inches
  )
  
  # US Congress Collage
  addWorksheet(wb, "US Congress Plots")
  
  temp_congress_plot <- tempfile(fileext = ".png")
  ggsave(temp_congress_plot, congress_plots, width = 12, height = rows * 3, units = "in", dpi = 700, limitsize = FALSE)
  
  insertImage(
    wb, 
    sheet = "US Congress Plots", 
    file = temp_congress_plot, 
    width = 12,  # Adjust size as needed
    height = rows*3,
    startRow = 1,  # Position in Excel (row 1, column 1)
    startCol = 1,
    units = "in"   # Dimensions in inches
  )
  
  
  # State House Collage
  addWorksheet(wb, "State House Plots")
  
  temp_state_house_plots <- tempfile(fileext = ".png")
  ggsave(temp_state_house_plots, state_house_plots, width = 12, height = rows * 3, units = "in", dpi = 700, limitsize = FALSE)
  
  insertImage(
    wb, 
    sheet = "State House Plots", 
    file = temp_state_house_plots, 
    width = 12,  # Adjust size as needed
    height = rows*3,
    startRow = 1,  # Position in Excel (row 1, column 1)
    startCol = 1,
    units = "in"   # Dimensions in inches
  )
  
  
  # Compiled Plot Collage
  addWorksheet(wb, "Compiled Plots")
  
  temp_compiled_plots <- tempfile(fileext = ".png")
  ggsave(temp_compiled_plots, compiled_plots, width = 12, height = rows * 3 + 2, units = "in", dpi = 700, limitsize = FALSE)
  
  insertImage(
    wb, 
    sheet = "Compiled Plots", 
    file = temp_compiled_plots, 
    width = 12,  # Adjust size as needed
    height = rows*3 + 2,
    startRow = 1,  # Position in Excel (row 1, column 1)
    startCol = 1,
    units = "in"   # Dimensions in inches
  )
  
  # Save the final Excel file
  saveWorkbook(wb, paste(county, ".xlsx", sep = ""), overwrite = TRUE)
  
  }
