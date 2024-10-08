# Load necessary libraries
library(ggplot2)
library(tidyverse)

# Example dataset
data <- data.frame(
  Category = c('A', 'B', 'C', 'D'),
  Values = c(23, 45, 12, 67)
)

# Create a plot
ggplot(data, aes(x = Category, y = Values, fill = Category)) +
  geom_bar(stat = "identity", width = 0.6, show.legend = FALSE) +
  labs(
    title = "Sample Canva-like Bar Plot",
    subtitle = "Aesthetic Design for Clean Visualizations",
    caption = "Source: Example Data"
  ) +
  theme_minimal(base_size = 15) +  # Use a minimal theme for the clean look
  theme(
    plot.title = element_text(face = "bold", size = 18, hjust = 0.5), # Centered title
    plot.subtitle = element_text(hjust = 0.5, color = "gray40"),       # Centered subtitle
    plot.caption = element_text(color = "gray60", size = 10),
    axis.title = element_blank(),     # Remove axis titles for simplicity
    axis.text = element_text(size = 12),
    axis.text.x = element_text(face = "bold"),
    panel.grid.major = element_blank(),   # Remove major gridlines
    panel.grid.minor = element_blank(),   # Remove minor gridlines
    panel.background = element_rect(fill = "white"),
    plot.background = element_rect(fill = "white"),
    legend.position = "none"             # No legend for simplicity
  ) +
  scale_fill_manual(values = c("#FF6F61", "#6B5B95", "#88B04B", "#F7CAC9"))  # Muted Canva-like color palette


# Example dataset (replace with your actual data)
data <- data.frame(
  Month = c("2023-07", "2023-08", "2023-09", "2023-10", "2023-11", "2023-12",
            "2024-01", "2024-02", "2024-03", "2024-04", "2024-05", "2024-06", "2024-07"),
  First_Visits = c(500, 600, 550, 580, 610, 670, 590, 550, 530, 570, 590, 550, 400),
  Return_Visits = c(2000, 2200, 2300, 2400, 2600, 2800, 2500, 2300, 2100, 2400, 2700, 2300, 1500)
)

# Reshape data for stacked bar plot
data_long <- data %>%
  pivot_longer(cols = c(First_Visits, Return_Visits),
               names_to = "Visit_Type", values_to = "Count")

# Create the plot
ggplot(data_long, aes(x = Month, y = Count, fill = Visit_Type)) +
  # Reduce the gap between bars and increase bar width
  geom_bar(stat = "identity", position = "stack", width = 0.8) +  # Adjusted width

  # Adding a rounded top effect using geom_bar and geom_rect
  geom_rect(aes(xmin = as.numeric(as.factor(Month)) - 2,
                xmax = as.numeric(as.factor(Month)) + 2,
                ymin = ifelse(Visit_Type == "First_Visits", 0, Count),
                ymax = Count),
            fill = NA, color = NA) +

  labs(
    title = "First vs Return",
    subtitle = "July 2023 - June 2024",
    x = NULL,  # Remove x-axis label for cleanliness
    y = NULL   # Remove y-axis label for cleanliness
  ) +

  # Scale for custom colors
  scale_fill_manual(
    values = c("First_Visits" = "#F28E2B", "Return_Visits" = "#4E79A7"),
    name = NULL,  # Remove legend title
    labels = c("First Visits", "Return Visits")
  ) +

  # Custom theme and style for the plot
  theme_minimal(base_size = 15) +  # Minimalist theme
  theme(
    plot.title = element_text(face = "bold", size = 20, hjust = 0),  # Title styling
    plot.subtitle = element_text(size = 14, hjust = 0),              # Subtitle styling
    axis.text.x = element_text(angle = 45, hjust = 1, vjust = 1),    # Rotate x-axis labels
    axis.text.y = element_text(size = 12),                           # Customize y-axis labels
    panel.grid.major.x = element_blank(),  # Remove vertical gridlines
    panel.grid.minor = element_blank(),    # Remove minor gridlines
    panel.grid.major.y = element_line(color = "lightgrey"),  # Keep horizontal grid lines light grey
    plot.background = element_rect(fill = "transparent", color = NA),    # Set transparent background
    panel.background = element_rect(fill = "transparent", color = NA),   # Set panel background to transparent
    legend.position = "top"  # Position legend at the top
  )
