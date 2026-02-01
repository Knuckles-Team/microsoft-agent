import asyncio
import os
from urllib.parse import urlparse, unquote

from crawl4ai import AsyncWebCrawler, CrawlerRunConfig
from crawl4ai.deep_crawling import BFSDeepCrawlStrategy, FilterChain, URLPatternFilter


async def main():
    # Starting URL
    start_url = "https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0&preserve-view=true"

    # Create output folder
    output_folder = "microsoft"
    os.makedirs(output_folder, exist_ok=True)

    # Filter to only include URLs in the Graph API docs section
    filter_chain = FilterChain(
        filters=[
            URLPatternFilter(patterns="https://learn.microsoft.com/en-us/graph/api/*")
        ]
    )

    # Configure deep crawl strategy: BFS up to depth 2, max 50 pages
    deep_strategy = BFSDeepCrawlStrategy(
        max_depth=2, max_pages=50, filter_chain=filter_chain
    )

    # Crawler config with deep strategy and filters
    config = CrawlerRunConfig(
        deep_crawl_strategy=deep_strategy,
        prefetch=False,  # Ensure full processing including Markdown
    )

    # Run the crawler
    async with AsyncWebCrawler(verbose=True) as crawler:  # verbose for progress logging
        results = await crawler.arun(url=start_url, config=config)

    # Save each page's Markdown to a file
    for result in results:
        if result.success and result.markdown:
            # Create a safe filename from the URL path (e.g., /graph/api/overview -> overview.md)
            parsed_url = urlparse(result.url)
            path_parts = parsed_url.path.strip("/").split("/")[2:]  # Skip 'en-us/graph'
            filename = "_".join(path_parts) or "index"
            filename = (
                unquote(filename).replace("?", "_").replace("&", "_") + ".md"
            )  # Sanitize
            filepath = os.path.join(output_folder, filename)

            with open(filepath, "w", encoding="utf-8") as f:
                f.write(result.markdown)

            print(f"Saved: {filepath} (from {result.url})")

    print(
        f"\nCrawl complete. {len(results)} pages processed and saved to '{output_folder}' folder."
    )


if __name__ == "__main__":
    asyncio.run(main())
