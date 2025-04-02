import datetime
import logging
import os
from contextlib import asynccontextmanager
from typing import AsyncGenerator, Dict, List, Optional, Tuple, Union

import numpy as np
import pywintypes
import win32com.client as win32
from PIL import Image

from .style import STYLE_CONFIGS, PPTStyle

formatter = logging.Formatter(
    fmt='%(asctime)s [%(levelname)s] %(name)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# Configure handlers with formatter
file_handler = logging.FileHandler("open-manus-ppt.log", encoding='utf-8')
file_handler.setFormatter(formatter)

console_handler = logging.StreamHandler()
console_handler.setFormatter(formatter)

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    handlers=[file_handler, console_handler]
)
logger = logging.getLogger("open-manus-ppt")

class PPTError(Exception):
    """Custom exception for PPT operations"""
    pass

def rgb_to_int(rgb_tuple: Tuple[int, int, int]) -> int:
    """Convert RGB tuple to PowerPoint color value"""
    return rgb_tuple[0] + (rgb_tuple[1] << 8) + (rgb_tuple[2] << 16)

class SlideManager:
    """PowerPoint slide manager class"""

    async def __aenter__(self) -> 'SlideManager':
        """Async context manager entry"""
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb) -> None:
        """Async context manager exit"""
        await self.cleanup()

    def __init__(self, ppt_path: str, style: PPTStyle):
        """
        Initialize SlideManager

        Args:
            ppt_path: Path to PowerPoint file
            style: PowerPoint style enum
        """
        self.style_config = STYLE_CONFIGS[style]
        self._ppt = None
        self._presentation = None
        self._new_file_flag = False
        self._save_path = ppt_path
        self._initialize(ppt_path)
        self._setup_master()

    def _initialize(self, ppt_path: str) -> None:
        """Initialize PowerPoint application and presentation"""
        try:
            self._ppt = win32.Dispatch('PowerPoint.Application')
            self._ppt.Visible = True
            logger.debug("PowerPoint application initialized")

            abs_path = os.path.abspath(ppt_path)
            if os.path.exists(abs_path):
                self._presentation = self._ppt.Presentations.Open(abs_path)
                logger.info(f"Opened existing file: {abs_path}")
            else:
                self._presentation = self._ppt.Presentations.Add()
                self._new_file_flag = True
                logger.info(f"Created new file: {abs_path}")
        except Exception as e:
            logger.error(f"Failed to initialize PowerPoint: {str(e)}")
            raise PPTError(f"PowerPoint initialization failed: {str(e)}")

    async def save(self) -> None:
        """Async save presentation"""
        try:
            if self._presentation:
                if self._new_file_flag:
                    self._presentation.SaveAs(self._save_path)
                    logger.info(f"New file saved at: {self._save_path}")
                else:
                    # self._presentation.Save()
                    self._presentation.SaveAs(self._save_path)
                    logger.info(f"File saved as: {self._save_path}")
        except Exception as e:
            logger.error(f"save failed: {str(e)}")
            raise PPTError(f"save failed: {str(e)}")
    async def cleanup(self) -> None:
        """Async cleanup of PowerPoint resources and save presentation"""
        try:
            await self.save()
            self._presentation = None

            if self._ppt:
                self._ppt.Quit()
                self._ppt = None
                logger.info("PowerPoint application closed")
        except Exception as e:
            logger.error(f"Cleanup failed: {str(e)}")
            raise PPTError(f"Cleanup failed: {str(e)}")


    def _setup_gradient_background(self, slide: 'win32.Dispatch') -> None:
        """Setup gradient background for slide"""
        try:
            # Get background
            background = slide.Background

            # Set fill type to gradient
            fill = background.Fill
            fill.ForeColor.RGB = rgb_to_int(self.style_config.theme_color["secondary"])
            fill.BackColor.RGB = rgb_to_int(self.style_config.theme_color["accent"])

            # Configure gradient
            fill.TwoColorGradient(1, 1)  # 1 = msoGradientHorizontal, 1 = Number of colors


            # Adjust transparency
            fill.Transparency = 0.1  # Slight transparency for better visual effect

            logger.debug("Gradient background configured")
        except Exception as e:
            logger.error(f"Failed to setup gradient background: {str(e)}")
            raise PPTError(f"Failed to setup gradient background: {str(e)}")
    def _setup_master(self) -> None:
        """Configure master slide style"""
        logger.debug("Configuring master style...")
        master = self._presentation.SlideMaster

        # Set theme colors
        for idx, (_, rgb) in enumerate([
            ("primary", "primary"),
            ("background", "background"),
            ("accent", "accent"),
            ("text_dark", "text_dark")
        ], 1):
            master.ColorScheme.Colors(idx).RGB = rgb_to_int(
                self.style_config.theme_color[rgb]
            )

        # Set default font
        master.Shapes.Title.TextFrame.TextRange.Font.Name = \
            self.style_config.fonts["title"]["name"]

        # Add gradient background
        self._setup_gradient_background(master)

        self._setup_footer(master)
        logger.debug("Master style configured")

    def _setup_footer(self, master: 'win32.Dispatch') -> None:
        """Setup footer"""
        footer = master.Shapes.AddShape(
            Type=1,
            Left=0,
            Top=self._presentation.PageSetup.SlideHeight - self.style_config.layout["footer_height"],
            Width=self._presentation.PageSetup.SlideWidth,
            Height=self.style_config.layout["footer_height"]
        )
        footer.Fill.Solid()
        footer.Fill.ForeColor.RGB = rgb_to_int(self.style_config.theme_color["secondary"])
        footer.Line.Visible = False

        self._add_page_number(master)

    def _add_page_number(self, master: 'win32.Dispatch') -> None:
        """Add page number to footer"""
        page_num = master.HeadersFooters.SlideNumber
        page_num.Visible = True
        num_box = master.Shapes.AddTextbox(
            Orientation=1,
            Left=self._presentation.PageSetup.SlideWidth/2 - 30,
            Top=self._presentation.PageSetup.SlideHeight - 40,
            Width=60,
            Height=30
        )
        num_box.TextFrame.TextRange.Text = "&p"
        num_box.TextFrame.TextRange.Font.Name = self.style_config.fonts["data"]["name"]
        num_box.TextFrame.TextRange.Font.Size = self.style_config.fonts["data"]["size"]

    def new_slide(self, layout_type: int = 7) -> int:
        """
        Add a new slide to the end of presentation

        Args:
            layout_type: PowerPoint layout type (default=2 for Title and Content)
                1 = Title Slide
                2 = Title and Content
                3 = Section Header
                4 = Two Content
                5 = Comparison
                6 = Title Only
                7 = Blank
                8 = Content with Caption
                9 = Picture with Caption

        Returns:
            int: Page number of the new slide

        Raises:
            PPTError: If failed to add new slide
        """
        try:
            # Add new slide at the end
            new_slide = self._presentation.Slides.Add(
                self._presentation.Slides.Count + 1,
                layout_type
            )
            page_num = new_slide.SlideIndex
            logger.info(f"Added new slide at position {page_num}")
            return page_num

        except Exception as e:
            logger.error(f"Failed to add new slide: {str(e)}")
            raise PPTError(f"Failed to add new slide: {str(e)}")
    def create_cover(self, title: str, subtitle: str = "", author: str = "") -> 'win32.Dispatch':
        """Create cover page"""
        logger.info(f"Creating cover page: {title}")
        slide = self._presentation.Slides.Add(1, 1)  # 1 = Title layout

        # Main title
        title_shape = slide.Shapes.Title
        title_shape.TextFrame.TextRange.Text = title
        title_shape.TextFrame.TextRange.Font.Name = self.style_config.fonts["title"]["name"]
        title_shape.TextFrame.TextRange.Font.Size = self.style_config.fonts["title"]["size"]
        title_shape.TextFrame.TextRange.Font.Color.RGB = rgb_to_int(
            self.style_config.theme_color["text_dark"]
        )

        # Subtitle and author with dark text
        if subtitle or author:
            sub_text = f"{subtitle}\n{author}" if subtitle and author else subtitle or author
            subtitle_shape = slide.Shapes.Placeholders(2)
            subtitle_shape.TextFrame.TextRange.Text = sub_text
            subtitle_shape.TextFrame.TextRange.Font.Name = self.style_config.fonts["subtitle"]["name"]
            subtitle_shape.TextFrame.TextRange.Font.Size = self.style_config.fonts["subtitle"]["size"]
            subtitle_shape.TextFrame.TextRange.Font.Color.RGB = rgb_to_int(
                self.style_config.theme_color["text_dark"]
            )

        return slide

    def create_table_of_content(self, items: List[str]):
        logger.info("Creating table of contents")
        slide = self._presentation.Slides.Add(self._presentation.Slides.Count + 1, 2)

        # Set title with dark text
        title_shape = slide.Shapes.Title
        title_shape.TextFrame.TextRange.Text = "Contents"
        title_shape.TextFrame.TextRange.Font.Name = self.style_config.fonts["title"]["name"]
        title_shape.TextFrame.TextRange.Font.Size = self.style_config.fonts["title"]["size"]
        title_shape.TextFrame.TextRange.Font.Color.RGB = rgb_to_int(
            self.style_config.theme_color["text_dark"]
        )

        # Add TOC items with dark text
        left_margin = self.style_config.layout["page_margin"]
        top = slide.Shapes.Title.Height + self.style_config.layout["grid_gutter"]

        for idx, text in enumerate(items):
            self._add_toc_item(slide, text, idx, left_margin,
                            top + idx * self.style_config.layout["grid_gutter"] * 2)


    def _add_toc_item(self, slide: 'win32.Dispatch', text: str,
                      number: int, left: float, top: float) -> None:
        """Add table of contents item"""
        # Number background
        num_shape = slide.Shapes.AddShape(1, left, top, 30, 30)
        num_shape.Fill.Solid()
        num_shape.Fill.ForeColor.RGB = rgb_to_int(self.style_config.theme_color["accent"])
        num_shape.TextFrame.TextRange.Text = str(number)
        num_shape.TextFrame.TextRange.Font.Color.RGB = rgb_to_int(
            self.style_config.theme_color["background"]
        )

        # TOC text
        text_box = slide.Shapes.AddTextbox(1, left + 40, top, 400, 30)
        text_box.TextFrame.TextRange.Text = f"{text}"
        text_box.TextFrame.TextRange.Font.Name = self.style_config.fonts["body"]["name"]
        text_box.TextFrame.TextRange.Font.Size = self.style_config.fonts["body"]["size"]
        text_box.TextFrame.TextRange.Font.Color.RGB = rgb_to_int(
            self.style_config.theme_color["text_dark"]
        )
    def goto_slide(self, page_num: int) -> None:
        """Navigate to specific slide"""
        if 1 <= page_num <= self._presentation.Slides.Count:
            self._presentation.Slides(page_num).Select()
            logger.info(f"Navigated to slide {page_num}")
        else:
            logger.error(f"Invalid page number: {page_num}")
            raise ValueError(f"Invalid page number: {page_num}")

    def add_line_chart(
        self,
        page_num: int,
        data: Dict[str, Union[List[str], List[Tuple[str, List[float]]]]],
        position: Tuple[float, float, float, float],
        title: str = "Line Chart",
        x_label: str = "Category",
        y_label: str = "Value"
    ) -> None:
        """Add multi-series line chart to specified slide

        Args:
            page_num: Target slide number
            data: Dictionary containing:
                - categories: List of x-axis labels
                - series: List of (series_name, values) pairs
            position: (left, top, width, height) of chart
            title: Chart title
            x_label: X-axis label
            y_label: Y-axis label
        """
        logger.info(f"Adding multi-series line chart to slide {page_num}")
        try:
            self.goto_slide(page_num)
            slide = self._presentation.Application.ActiveWindow.View.Slide

            # Add chart
            left, top, width, height = position
            chart = slide.Shapes.AddChart2(-1, 4).Chart  # 4 = Line chart

            # Set position and size
            chart.Parent.Left = left
            chart.Parent.Top = top
            chart.Parent.Width = width
            chart.Parent.Height = height

            # Set data
            worksheet = chart.ChartData.Workbook.Worksheets(1)
            categories = data["categories"]
            series_data = data["series"]

            # Write categories (x-axis labels)
            worksheet.Cells(1, 1).Value = x_label
            for i, category in enumerate(categories, start=2):
                worksheet.Cells(i, 1).Value = category

            # Write series data
            for col, (series_name, values) in enumerate(series_data, start=2):
                worksheet.Cells(1, col).Value = series_name
                for row, value in enumerate(values, start=2):
                    worksheet.Cells(row, col).Value = value

            # Configure chart title with dark text
            chart.HasTitle = True
            chart.ChartTitle.Text = title
            title_font = chart.ChartTitle.Characters.Font
            title_font.Color = rgb_to_int(self.style_config.theme_color["text_dark"])

            # Set axis labels to dark text
            for axis in [chart.Axes(1), chart.Axes(2)]:  # 1 = x-axis, 2 = y-axis
                axis.HasTitle = True
                axis.AxisTitle.Text = x_label if axis.Type == 1 else y_label
                axis_font = axis.TickLabels.Font
                axis_font.Color = rgb_to_int(self.style_config.theme_color["text_dark"])
                axis.AxisTitle.Font.Color = rgb_to_int(self.style_config.theme_color["text_dark"])

            # Set modern style
            chart.ChartStyle = 34

            chart.ChartData.Workbook.Close()
            logger.info("Multi-series line chart added successfully")
        except Exception as e:
            logger.error(f"Failed to add line chart: {str(e)}")
            raise PPTError(f"Failed to add line chart: {str(e)}")

    def add_treemap(
        self,
        page_num: int,
        data: Dict[str, Union[float, Dict[str, float]]],
        position: Tuple[float, float, float, float],
        title: str = "Treemap Chart"
    ) -> None:
        """
        Add treemap chart to specified slide

        Args:
            page_num: Target slide number
            data: Nested dictionary of categories and values
            position: (left, top, width, height) of chart
            title: Chart title
        """
        logger.info(f"Adding treemap to slide {page_num}")
        try:
            self.goto_slide(page_num)
            slide = self._presentation.Application.ActiveWindow.View.Slide

            # Add chart
            left, top, width, height = position
            chart = slide.Shapes.AddChart2(-1, 117).Chart  # 117 = Treemap

            # Set position and size
            chart.Parent.Left = left
            chart.Parent.Top = top
            chart.Parent.Width = width
            chart.Parent.Height = height

            # Set data
            worksheet = chart.ChartData.Workbook.Worksheets(1)
            self._populate_treemap_data(worksheet, data)

            # Configure chart
            chart.HasTitle = True
            chart.ChartTitle.Text = title
            chart.ChartStyle = 34  # Modern style

            chart.ChartData.Workbook.Close()
            logger.info("Treemap added successfully")
        except Exception as e:
            logger.error(f"Failed to add treemap: {str(e)}")
            raise PPTError(f"Failed to add treemap: {str(e)}")

    def add_radar_chart(
        self,
        page_num: int,
        data: List[Tuple[str, List[float]]],
        categories: List[str],
        position: Tuple[float, float, float, float],
        title: str = "Radar Chart"
    ) -> None:
        """
        Add radar chart to specified slide

        Args:
            page_num: Target slide number
            data: List of (series_name, values) pairs
            categories: List of category names
            position: (left, top, width, height) of chart
            title: Chart title
        """
        logger.info(f"Adding radar chart to slide {page_num}")
        try:
            self.goto_slide(page_num)
            slide = self._presentation.Application.ActiveWindow.View.Slide

            # Add chart
            left, top, width, height = position
            chart = slide.Shapes.AddChart2(-1, 76).Chart  # 76 = Radar

            # Set position and size
            chart.Parent.Left = left
            chart.Parent.Top = top
            chart.Parent.Width = width
            chart.Parent.Height = height

            # Set title with dark text
            chart.HasTitle = True
            chart.ChartTitle.Text = title
            title_font = chart.ChartTitle.Characters.Font
            title_font.Color = rgb_to_int(self.style_config.theme_color["text_dark"])

            # Set axis labels to dark text
            for axis in [chart.Axes(1), chart.Axes(2)]:  # 1 = x-axis, 2 = y-axis
                axis_font = axis.TickLabels.Font
                axis_font.Color = rgb_to_int(self.style_config.theme_color["text_dark"])

            # Set data
            worksheet = chart.ChartData.Workbook.Worksheets(1)

            # Add categories
            for i, category in enumerate(categories, start=2):
                worksheet.Cells(i, 1).Value = category

            # Add series data
            for col, (series_name, values) in enumerate(data, start=2):
                worksheet.Cells(1, col).Value = series_name
                for row, value in enumerate(values, start=2):
                    worksheet.Cells(row, col).Value = value

            # Configure chart
            chart.HasTitle = True
            chart.ChartTitle.Text = title
            chart.ChartStyle = 34  # Modern style

            chart.ChartData.Workbook.Close()
            logger.info("Radar chart added successfully")
        except Exception as e:
            logger.error(f"Failed to add radar chart: {str(e)}")
            raise PPTError(f"Failed to add radar chart: {str(e)}")

    def _convert_to_com_date(self, date_obj: datetime.date) -> pywintypes.TimeType:
        """Convert Python date to COM date"""
        return pywintypes.Time(datetime.datetime.combine(date_obj, datetime.time()))
    def add_gantt_chart(
        self,
        page_num: int,
        tasks: List[Tuple[str, datetime.date, datetime.date]],
        position: Tuple[float, float, float, float],
        title: str = "Project Timeline"
    ) -> None:
        """
        Add Gantt chart to specified slide

        Args:
            page_num: Target slide number
            tasks: List of (task_name, start_date, end_date) tuples
            position: (left, top, width, height) of chart
            title: Chart title
        """
        logger.info(f"Adding Gantt chart to slide {page_num}")
        try:
            self.goto_slide(page_num)
            slide = self._presentation.Application.ActiveWindow.View.Slide

            # Add chart
            left, top, width, height = position
            chart = slide.Shapes.AddChart2(-1, 24).Chart  # 24 = Bar chart

            # Set position and size
            chart.Parent.Left = left
            chart.Parent.Top = top
            chart.Parent.Width = width
            chart.Parent.Height = height

            # Set title with dark text
            chart.HasTitle = True
            chart.ChartTitle.Text = title
            title_font = chart.ChartTitle.Characters.Font
            title_font.Color = rgb_to_int(self.style_config.theme_color["text_dark"])

            # Set axis labels to dark text
            for axis in [chart.Axes(1), chart.Axes(2)]:  # 1 = x-axis, 2 = y-axis
                axis_font = axis.TickLabels.Font
                axis_font.Color = rgb_to_int(self.style_config.theme_color["text_dark"])

            # Set data
            worksheet = chart.ChartData.Workbook.Worksheets(1)
            worksheet.Cells(1, 1).Value = "Task"
            worksheet.Cells(1, 2).Value = "Start"
            worksheet.Cells(1, 3).Value = "Duration"

            for i, (task, start, end) in enumerate(tasks, start=2):
                # Convert dates to COM-compatible format
                com_start = self._convert_to_com_date(start)
                worksheet.Cells(i, 1).Value = task
                worksheet.Cells(i, 2).Value = com_start
                duration = (end - start).days
                worksheet.Cells(i, 3).Value = duration

            # Configure chart
            chart.ChartStyle = 34  # Modern style

            # Customize for Gantt appearance
            chart.PlotBy = 2  # By rows
            chart.Axes(2).MaximumScale = self._convert_to_com_date(max(t[2] for t in tasks))
            chart.Axes(2).MinimumScale = self._convert_to_com_date(min(t[1] for t in tasks))

            chart.ChartData.Workbook.Close()
            logger.info("Gantt chart added successfully")
        except Exception as e:
            logger.error(f"Failed to add Gantt chart: {str(e)}")
            raise PPTError(f"Failed to add Gantt chart: {str(e)}")

    # Add this method to the SlideManager class
    def add_text(
        self,
        page_num: int,
        text: str,
        position: Tuple[float, float, float, float],
        font_size: Optional[int] = None,
        font_name: Optional[str] = None,
        is_title: bool = False,
        is_bullet: bool = False
    ) -> None:
        """
        Add text to specified slide

        Args:
            page_num: Target slide number
            text: Text content to add
            position: (left, top, width, height) of text box
            font_size: Custom font size (optional)
            font_name: Custom font name (optional)
            is_title: Whether this is a title text
            is_bullet: Whether to add bullet points

        Raises:
            PPTError: If failed to add text
        """
        logger.info(f"Adding text to slide {page_num}")
        try:
            self.goto_slide(page_num)
            slide = self._presentation.Application.ActiveWindow.View.Slide

            # Create text box
            left, top, width, height = position
            text_box = slide.Shapes.AddTextbox(
                Orientation=1,  # msoTextOrientationHorizontal
                Left=left,
                Top=top,
                Width=width,
                Height=height
            )

            # Configure text frame
            text_frame = text_box.TextFrame
            text_frame.WordWrap = True
            text_frame.AutoSize = 1  # msoAutoSizeShapeToFitText

            # Set text
            text_range = text_frame.TextRange
            text_range.Text = text

            # Apply font settings
            font = text_range.Font
            if is_title:
                font.Name = font_name or self.style_config.fonts["title"]["name"]
                font.Size = font_size or self.style_config.fonts["title"]["size"]
                font.Bold = self.style_config.fonts["title"].get("bold", False)
            else:
                font.Name = font_name or self.style_config.fonts["body"]["name"]
                font.Size = font_size or self.style_config.fonts["body"]["size"]
                font.Bold = self.style_config.fonts["body"].get("bold", False)

            # Set text color
            font.Color.RGB = rgb_to_int(self.style_config.theme_color["text_dark"])

            # Add bullet points if requested
            if is_bullet:
                paragraphs = text_range.Paragraphs()
                for para in paragraphs:
                    para.ParagraphFormat.Bullet.Type = 1  # msoBullet
                    para.ParagraphFormat.Bullet.Character = 8226  # bullet character "•"

            logger.info("Text added successfully")
        except Exception as e:
            logger.error(f"Failed to add text: {str(e)}")
            raise PPTError(f"Failed to add text: {str(e)}")

    def add_image(
        self,
        page_num: int,
        image_path: str,
        position: Tuple[float, float, float, float],
        target_dpi: int = 300
    ) -> None:
        """
        Add image to specified slide with resolution adjustment

        Args:
            page_num: Target slide number
            image_path: Path to image file
            position: (left, top, width, height) of image
            target_dpi: Target DPI for image
        """
        logger.info(f"Adding image to slide {page_num}")
        try:
            # Process image
            with Image.open(image_path) as img:
                # Calculate new size based on target DPI
                current_dpi = img.info.get('dpi', (72, 72))[0]
                scale_factor = target_dpi / current_dpi
                new_size = (
                    int(img.width * scale_factor),
                    int(img.height * scale_factor)
                )

                # Resize image
                resized_img = img.resize(new_size, Image.LANCZOS)

                # Save temporary file
                temp_path = f"{image_path}_temp.png"
                resized_img.save(temp_path, dpi=(target_dpi, target_dpi))

            # Add to slide
            self.goto_slide(page_num)
            slide = self._presentation.Application.ActiveWindow.View.Slide

            left, top, width, height = position
            picture = slide.Shapes.AddPicture(
                os.path.abspath(temp_path),
                LinkToFile=False,
                SaveWithDocument=True,
                Left=left,
                Top=top,
                Width=width,
                Height=height
            )

            # Cleanup
            os.remove(temp_path)
            logger.info("Image added successfully")

        except Exception as e:
            logger.error(f"Failed to add image: {str(e)}")
            if os.path.exists(temp_path):
                os.remove(temp_path)
            raise PPTError(f"Failed to add image: {str(e)}")

@asynccontextmanager
async def slide_manager(ppt_path: str, style: PPTStyle = PPTStyle.FRESH_GREEN) -> AsyncGenerator[SlideManager, None]:
    """
    Async PowerPoint slide context manager

    Args:
        ppt_path: Path to PowerPoint file
        style: PowerPoint style enum

    Yields:
        SlideManager instance
    """
    logger.info(f"Initializing PPT manager: {ppt_path}, Style: {style.value}")
    manager = None

    try:
        manager = SlideManager(ppt_path, style)
        async with manager as slide_mgr:
            yield slide_mgr
    except Exception as e:
        logger.error(f"PPT manager error: {str(e)}", exc_info=True)
        raise PPTError(f"PPT manager error: {str(e)}")
