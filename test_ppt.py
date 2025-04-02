import asyncio
import datetime
import os

from app.tool.ppt import PPTStyle, slide_manager


async def amain():
    """Demonstrate the usage of PPT manager"""
    output_path = os.path.join(os.path.expanduser("~"), "Documents", "demo.pptx")

    async with slide_manager(output_path, style=PPTStyle.FRESH_GREEN) as manager:
        # Create cover page
        manager.create_cover(
            title="2025 Strategic Planning",
            subtitle="Digital Transformation Journey",
            author="Strategic Planning Department"
        )

        # Create table of contents
        toc_items = [
            "Market Analysis",
            "Competitive Landscape",
            "Strategic Goals",
            "Implementation Path",
            "Timeline",
        ]
        manager.create_table_of_content(toc_items)


        n = manager.new_slide()
        # Add line chart
        line_data = [
            ("Jan", 100), ("Feb", 120), ("Mar", 140),
            ("Apr", 130), ("May", 150)
        ]
        manager.add_line_chart(
            page_num=n,
            data=line_data,
            position=(100, 100, 600, 400),
            title="Monthly Sales Growth"
        )
        manager.add_text(
            page_num=n,
            text="This is a sample line chart title.",
            position=(100, 450, 600, 40),
            font_size=14,
            font_name="Arial",
            is_title=True
        )
        manager.add_text(
            page_num=n,
            text="This is a sample line chart text.",
            position=(100, 490, 600, 50),
            is_title=False
        )

        # Add radar chart
        n = manager.new_slide()
        radar_data = [
            ("Company A", [80, 90, 70, 85, 75]),
            ("Company B", [85, 80, 90, 70, 80])
        ]
        categories = ["Quality", "Price", "Service", "Innovation", "Brand"]
        manager.add_radar_chart(
            page_num=n,
            data=radar_data,
            categories=categories,
            position=(100, 100, 600, 400),
            title="Competitive Analysis"
        )

        # Add Gantt chart
        n = manager.new_slide()
        today = datetime.date.today()
        tasks = [
            ("Planning", today, today + datetime.timedelta(days=30)),
            ("Development", today + datetime.timedelta(days=30),
             today + datetime.timedelta(days=90)),
            ("Testing", today + datetime.timedelta(days=90),
             today + datetime.timedelta(days=120))
        ]
        manager.add_gantt_chart(
            page_num=n,
            tasks=tasks,
            position=(100, 100, 600, 400),
            title="Project Timeline"
        )

        # Add image
        n = manager.new_slide()
        img_path = os.path.join(os.path.expanduser("~"), "Documents", "img.jpg")
        manager.add_image(
            page_num=n,
            image_path=img_path,
            position=(100, 100, 400, 300),
            target_dpi=300
        )

if __name__ == "__main__":
    asyncio.run(amain())
