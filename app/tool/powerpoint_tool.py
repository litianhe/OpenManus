import asyncio
import os
from datetime import datetime, timedelta
from typing import Dict, Generic, List, Optional, Tuple, TypeVar, Union

from pydantic import Field, field_validator
from pydantic_core.core_schema import ValidationInfo

from app.tool.base import BaseTool, ToolResult
from app.tool.ppt import PPTStyle, SlideManager, slide_manager

_PPT_DESCRIPTION = """\
PowerPoint presentations Creation tool, which can create presentations using ACTIONS mentioned below.
It is designed to create presentations following these guidelines.

# Input Format
Requirements for the presentation

# Response Rules
1. RESPONSE FORMAT: You must ALWAYS respond with valid JSON in this specified format, refer to examples below:
example#1: {'action': 'create_presentation', 'output_name': 'x.pptx'}
example#2: {'action': 'save_presentation'}

2. ACTIONS: You can use the following actions to create the presentation, and one action for one function calling.
Basic Operations:
- Create new presentation with function 'create_presentation':
{
    "create_presentation": {
        "output_name": "presentation.pptx"
    }
}


Slide Operations:
- Create cover slide with function 'create_cover_slide':
{
    "create_cover_slide": {
        "title": "Digital Transformation Strategy",
        "subtitle": "2025 Vision and Implementation",
        "author": "Strategic Planning Team"
    }
}

- Create table of contents with function 'create_table_of_content_slide':
{
    "create_table_of_content_slide": {
        "content": "Market Analysis\nCompetitive Landscape\nStrategic Goals\nImplementation Plan\nTimeline"
    }
}


- Create a empty slide with function 'create_content_slide':
{
    "create_content_slide": {}
}


- Line chart with explanation with function 'add_line_chart':
{
  "add_line_chart": {
    "slide_number": 3,
    "data": {
      "categories": ["2023 Q1", "2023 Q2", "2023 Q3", "2023 Q4", "2024 Q1"],
      "series": [
        ["Gold", [1800, 1850, 1900, 1950, 2000]],
        ["Silver", [24, 25, 26, 27, 28]]
      ]
    },
    "position": [100, 100, 600, 400],
    "title": "Precious Metals Price Trends",
    "x_label": "Quarter",
    "y_label": "Price (USD)"
  }
}

- Radar chart for comparison with function 'add_radar_chart':
{
  "add_radar_chart": {
    "slide_number": 4,
    "radar_chart_data": {
      "series": [
        ["Team Alpha", [95, 88, 92, 85, 90, 87]],
        ["Team Beta", [82, 95, 88, 90, 85, 83]],
        ["Team Gamma", [88, 84, 95, 92, 80, 89]]
      ],
      "categories": [
        "Problem Solving",
        "Technical Skills",
        "Communication",
        "Collaboration",
        "Innovation",
        "Project Management"
      ]
    },
    "position": [100, 100, 600, 400],
    "title": "Team Capabilities Assessment"
  }
}
- Gantt chart for timeline wtih function 'add_gantt_chart':
{
  "add_gantt_chart": {
    "slide_number": 5,
    "gantt_chart_data": [
      ["Requirements Analysis", "2024-01-01", "2024-01-15"],
      ["System Design", "2024-01-15", "2024-02-15"],
      ["Frontend Development", "2024-02-15", "2024-04-01"],
      ["Backend Development", "2024-02-15", "2024-04-01"],
      ["Integration Testing", "2024-04-01", "2024-04-15"],
      ["User Acceptance Testing", "2024-04-15", "2024-05-01"],
      ["Deployment Preparation", "2024-05-01", "2024-05-15"],
      ["Go Live", "2024-05-15", "2024-05-16"]
    ],
    "position": [50, 50, 860, 440],
    "title": "Software Development Project Timeline 2024"
  }
}

- Add formatted text with function 'add_text':
{
  "add_text": {
    "slide_number": 2,
    "content": [
      "Strategic Objectives 2024-2025",
      "• Increase market share by 25% in key regions\n• Launch 3 new product lines\n• Achieve 95% customer satisfaction score\n• Expand into 5 new international markets",
      "Implementation Timeline:\nPhase 1: Market Research (Q1 2024)\nPhase 2: Product Development (Q2-Q3 2024)\nPhase 3: Market Launch (Q4 2024)"
    ],
    "position": [
      [50, 50, 860, 80],
      [100, 150, 760, 200],
      [100, 300, 760, 200]
    ],
    "font_size": [32, 24, 20],
    "font_name": "Arial",
    "is_title": [true, false, false],
    "is_bullet": [false, true, false]
  }
}

Final Operations:
- Save (for completing the entire task) with function 'save_presentation':
{
    "save_presentation": {}
}

Error Cases:
- Save on error (when encountering issues):
{
    "save_presentation": {}
}

Remember:
- Always save the presentation when finished
- Use save_presentation when switching between major sections


3. TASK COMPLETION:
- Use "save_presentation" action once all the sliders are created.



Slide Layout Guidelines:
- Dimensions: 960x540 points (16:9)
- Safe margins: left/right: 50pt, top/bottom: 20pt
- Charts: position=[100,100,600,400]
- Text and Charts should be well organized and not overlap

For presentation creation:
- New presentation: create_presentation with output_name="..." and style="..."
- Cover slide: create_cover_slide with title="...", subtitle="..."
- Table of contents slide: create_table_of_content_slide with string separated by newlines
- Charts: add_line_chart/add_radar_chart/add_gantt_chart with appropriate data
- Text: add_text with slide_number=N, content="...", position=[x,y,w,h]
- New slide: create_content_slide is used to create blank slide and return the slide number
- Save: save_presentation to save the presentation
- Ensure formatting is consistent
- Ensure all charts and text box positions are within safe margins

Consider:
1. Logical flow and structure
2. Visual hierarchy
3. Consistent formatting
4. Balance of content
5. Appropriate chart usage

Track progress and maintain presentation quality.
"""

Context = TypeVar("Context")

class PPTCreationTool(BaseTool, Generic[Context]):
    name: str = "ppt_creation_tool"
    """Tool for creating PowerPoint presentations."""
    description: str = _PPT_DESCRIPTION
    parameters: dict = {
        "type": "object",
        "properties": {
            "action": {
                "type": "string",
                "enum": [
                    "create_presentation",
                    "save_presentation",
                    "create_presentation",
                    "save_presentation",
                    "create_cover_slide",
                    "create_table_of_content_slide",
                    "create_content_slide",
                    "add_line_chart",
                    "add_radar_chart",
                    "add_gantt_chart",
                    "add_text",
                ],
                "description": "The PowerPoint action to perform",
            },
            "output_name": {
                "type": "string",
                "description": "The filename which will be pass to create_presentation function, example: 'demo.pptx'",
            },
"style": {
                "type": "string",
                "enum": ["FRESH_GREEN", "BUSINESS_BLUE"],
                "description": "Presentation style theme",
            },
            "title": {
                "type": "string",
                "description": "Title for slides or charts",
            },
            "subtitle": {
                "type": "string",
                "description": "Subtitle for cover slide",
            },
            "content": {
                "type": "string",
                "description": "Content text for slides or TOC items (separated by newlines)",
            },
            "slide_number": {
                "type": "integer",
                "description": "Slide's page number, used by function 'add_line_chart' or 'add_text' or 'add_radar_chart' or 'add_gantt_chart' , if not specified, it will create a new last slide",
            },
            "position": {
                "type": "array",
                "items": {"type": "number"},
                "description": "Position coordinates [left, Top, width, height],  left is x-coordinate , Top is y-coordinate. Slide dimension is 16:9 aspect ratio 960x540 points",
            },
            "line_chart_data": {
                "type": "object",
                "properties": {
                    "categories": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "X-axis labels"
                    },
                    "series": {
                        "type": "array",
                        "items": {
                            "type": "array",
                            "items": [
                                {"type": "string"},
                                {
                                    "type": "array",
                                    "items": {"type": "number"}
                                }
                            ]
                        },
                        "description": "Series data: [[series_name, values], ...]"
                    }
                },
                "required": ["categories", "series"],
                "description": "Data structure for multi-series line chart"
            },
            "radar_chart_data": {
                "type": "object",
                "properties": {
                    "series": {
                        "type": "array",
                        "items": {
                            "type": "array",
                            "items": [{"type": "string"}, {"type": "array", "items": {"type": "number"}}]
                        },
                        "description": "Series data: [[name, values], ...]"
                    },
                    "categories": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Category labels"
                    }
                }
            },
            "gantt_chart_data": {
                "type": "array",
                "items": {
                    "type": "array",
                    "items": [
                        {"type": "string"},
                        {"type": "string", "format": "date"},
                        {"type": "string", "format": "date"}
                    ]
                },
                "description": "Tasks for Gantt chart: [[task_name, start_date, end_date], ...]"
            }
        },
        "required": ["action"],
        "dependencies": {
            "create_presentation": ["output_name"],
            "save_presentation": [],
            "create_cover_slide": ["title", "content"],
            "create_table_of_content_slide": ["content"],
            "create_content_slide": [],
            "add_line_chart": ["line_chart_data"],
            "add_radar_chart": ["radar_chart_data"],
            "add_gantt_chart": ["gantt_chart_data"],
            "add_text": ["slide_number", "content"],
        }
    }

    lock: asyncio.Lock = Field(default_factory=asyncio.Lock)
    manager: Optional[SlideManager] = Field(default=None, exclude=True)
    output_name: Optional[str] = Field(default=None)
    current_style: Optional[PPTStyle] = Field(default=None)
    output_path: Optional[str] = Field(default=None)

    @field_validator("parameters", mode="before")
    def validate_parameters(cls, v: dict, info: ValidationInfo) -> dict:
        if not v:
            raise ValueError("Parameters cannot be empty")
        return v

    async def _ensure_initialized(self, output_name: str, style: PPTStyle = PPTStyle.FRESH_GREEN):
        """Ensure PPT manager is initialized."""
        if self.manager is None or output_name != self.output_name:
            if self.manager:
                await self.cleanup()
            self.output_name = output_name
            self.current_style = style
            self.output_path = os.path.join(os.path.expanduser("~"), "Documents", output_name)
            # self.manager = await slide_manager(self.output_path, style=style).__aenter__()
            self.manager = SlideManager(self.output_path, style)
        return self.manager

    async def execute(
        self,
        action: str,
        output_name: Optional[str] = None,
        style: Optional[str] = None,
        title: Optional[str] = None,
        subtitle: Optional[str] = None,
        content: Optional[str] = None,
        image_path: Optional[str] = None,
        slide_number: Optional[int] = None,
        position: Optional[List[float]] = None,
        line_chart_data: Optional[Dict[str, Union[List[str], List[Tuple[str, List[float]]]]]] = None,
        radar_chart_data: Optional[Dict[str, Union[List[Tuple[str, List[float]]], List[str]]]] = None,
        gantt_chart_data: Optional[List[Tuple[str, str, str]]] = None,
        **kwargs,
    ) -> ToolResult:
        """Execute a specified PowerPoint action."""
        async with self.lock:
            try:
                # Initialize manager if needed
                if action == "create_presentation":
                    if not output_name:
                        return ToolResult(error="Output file name is required for creating presentation")
                    ppt_style = PPTStyle[style] if style else PPTStyle.FRESH_GREEN
                    await self._ensure_initialized(output_name, ppt_style)
                    return ToolResult(output=f"Created new presentation with {ppt_style.value} style")

                if not self.manager:
                    return ToolResult(error="Please create presentation first using 'create_presentation' action")

                # Handle different actions
                if action == "create_cover_slide":
                    if not title:
                        return ToolResult(error="Title is required for cover slide")
                    self.manager.create_cover(
                        title=title,
                        subtitle=subtitle or "",
                        author=kwargs.get("author", "")
                    )
                    return ToolResult(output="Created cover slide")

                elif action == "create_table_of_content_slide":
                    if not isinstance(content, str):
                        return ToolResult(error="Content must be a string with items separated by newlines")
                    toc_items = [item.strip() for item in content.split('\n') if item.strip()]
                    self.manager.create_table_of_content(toc_items)
                    return ToolResult(output="Created table of contents")

                elif action == "create_content_slide":
                    slide_num = self.manager.new_slide()
                    return ToolResult(output=f"Created a new slide with page number {slide_num}")

                elif action == "add_line_chart":
                    if not line_chart_data or "categories" not in line_chart_data or "series" not in line_chart_data:
                        return ToolResult(error="Line chart data with categories and series is required")

                    slide_num = slide_number or self.manager.new_slide()
                    pos = position or [100, 100, 600, 400]

                    # Extract categories and series from the new data structure
                    categories = line_chart_data["categories"]
                    series_data = line_chart_data["series"]

                    self.manager.add_line_chart(
                        page_num=slide_num,
                        data={
                            "categories": categories,
                            "series": series_data
                        },
                        position=tuple(pos),
                        title=title or "Line Chart",
                        x_label=kwargs.get("x_label", "Category"),
                        y_label=kwargs.get("y_label", "Value")
                    )
                    return ToolResult(output="Added multi-series line chart")

                elif action == "add_radar_chart":
                    if not radar_chart_data or "series" not in radar_chart_data or "categories" not in radar_chart_data:
                        return ToolResult(error="Radar chart data and categories are required")
                    slide_num = slide_number or self.manager.new_slide()
                    pos = position or [100, 100, 600, 400]
                    self.manager.add_radar_chart(
                        page_num=slide_num,
                        data=radar_chart_data["series"],
                        categories=radar_chart_data["categories"],
                        position=tuple(pos),
                        title=title or "Radar Chart"
                    )
                    return ToolResult(output="Added radar chart")

                elif action == "add_gantt_chart":
                    if not gantt_chart_data:
                        return ToolResult(error="Gantt chart tasks are required")
                    slide_num = slide_number or self.manager.new_slide()
                    pos = position or [100, 100, 600, 400]

                    # Convert string dates to datetime objects
                    tasks = [(name, datetime.strptime(start, "%Y-%m-%d").date(),
                             datetime.strptime(end, "%Y-%m-%d").date())
                            for name, start, end in gantt_chart_data]

                    self.manager.add_gantt_chart(
                        page_num=slide_num,
                        tasks=tasks,
                        position=tuple(pos),
                        title=title or "Project Timeline"
                    )
                    return ToolResult(output="Added Gantt chart")

                elif action == "create_image_slide":
                    if not image_path:
                        return ToolResult(error="Image path is required")
                    slide_num = slide_number or self.manager.new_slide()
                    pos = position or [100, 100, 400, 300]
                    self.manager.add_image(
                        page_num=slide_num,
                        image_path=image_path,
                        position=tuple(pos),
                        target_dpi=300
                    )
                    return ToolResult(output="Created image slide")

                elif action == "add_text":
                    slide_num = slide_number or self.manager.new_slide()
                    if not content:
                        return ToolResult(error="Content are required")
                    pos = position or [100, 100, 400, 200]
                    self.manager.add_text(
                        page_num=slide_number,
                        text=content,
                        position=tuple(pos),
                        font_size=kwargs.get("font_size"),
                        font_name=kwargs.get("font_name"),
                        is_title=kwargs.get("is_title", False),
                        is_bullet=kwargs.get("is_bullet", False)
                    )
                    return ToolResult(output="Added text to slide")

                elif action == "save_presentation":
                    await self.save()
                    return ToolResult(output=f"Saved presentation to {self.output_name}")

                else:
                    return ToolResult(error=f"Unknown action: {action}")

            except Exception as e:
                return ToolResult(error=f"PowerPoint action '{action}' failed: {str(e)}")

    async def save(self):
        """Clean up PPT manager resources."""
        if self.manager is not None:
            await self.manager.save()

    def __del__(self):
        """Ensure cleanup when object is destroyed."""
        # if self.manager is not None:
        #     try:
        #         asyncio.run(self.cleanup())
        #     except RuntimeError:
        #         loop = asyncio.new_event_loop()
        #         loop.run_until_complete(self.cleanup())
        #         loop.close()
        pass

