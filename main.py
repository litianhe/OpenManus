import asyncio

from app.agent.manus import Manus
from app.logger import logger

PROMPT = """
Create a PowerPoint presentation named "showcase.pptx" to demonstrate AI development capabilities, with the following requirements:

1. Cover Page
- Title: "AI Development Showcase"
- Subtitle: "2023-2024 Market Analysis"
- Author: "AI Development Showcase Team"

2. Table of Contents Page(Agenda)
- Hardware
- Infrastructure
- LLMs
- Agents
- Apps
- Tools

3. Market Analysis Page (Line Chart)
- Title: "Global AI Market"
- Chart data showing quarterly revenue growth line chart for 2024 , from AI Agent, LLM, and AI Platform:
- X-axis: "Quarters"
- Y-axis: "Market Size (Million USD)"

4.Layer comparison Page (Line Chart)
- Title: "AI Development Layers"
- Chart data showing quarterly revenue growth line chart for 2024, from AI Agent, LLM, and AI Platform:

5. Radar Chart Page
- Title: "LLM Analysis"
- Chart data compare ChatGPT/Deepseek/Qwen: Tech vs Revenue vs Regulation

6. Gantt Chart Page (Project Timeline)
- Title: "OpenAI/Meta/Tesla milestones"


7. Summary Page
- Key trends: Chip-model synergy | OSS growth | Regulatory gaps


Requirements:
- Use FRESH_GREEN theme
- Consistent formatting
- Professional layout
- Clear data visualization
- slider Format: 16:9 PPTX
- Fashion Icons wtih Tech-related symbols
- You can create some damo data for Charts

Please generate this presentation using available tools
"""

# test with config in config.toml as following:
# [llm]
# model = "Qwen/QwQ-32B"                      # The LLM model to use ,32k
# base_url = "https://api.siliconflow.cn/v1"  # API endpoint URL
# api_key = "sk-xxxx"                         # Your API key
# max_tokens = 8192                           # Maximum number of tokens in the response
# temperature = 0.0                           # Controls randomness


async def main():
    agent = Manus()
    try:
        # prompt = input("Enter your prompt: ")
        prompt = PROMPT
        if not prompt.strip():
            logger.warning("Empty prompt provided.")
            return

        logger.warning("Processing your request...")
        await agent.run(prompt)
        logger.info("Request processing completed.")
    except KeyboardInterrupt:
        logger.warning("Operation interrupted.")


if __name__ == "__main__":
    asyncio.run(main())
