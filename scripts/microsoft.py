import asyncio
import os
from urllib.parse import urlparse, unquote

from crawl4ai import AsyncWebCrawler, CrawlerRunConfig
from crawl4ai.deep_crawling import BFSDeepCrawlStrategy, FilterChain, URLPatternFilter


async def main():
    start_url = "https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0&preserve-view=true"

    output_folder = "microsoft"
    os.makedirs(output_folder, exist_ok=True)

    filter_chain = FilterChain(
        filters=[
            URLPatternFilter(patterns="https://learn.microsoft.com/en-us/graph/api/*")
        ]
    )

    deep_strategy = BFSDeepCrawlStrategy(
        max_depth=2, max_pages=50, filter_chain=filter_chain
    )

    config = CrawlerRunConfig(
        deep_crawl_strategy=deep_strategy,
        prefetch=False,
    )

    async with AsyncWebCrawler(verbose=True) as crawler:
        results = await crawler.arun(url=start_url, config=config)

    for result in results:
        if result.success and result.markdown:
            parsed_url = urlparse(result.url)
            path_parts = parsed_url.path.strip("/").split("/")[2:]
            filename = "_".join(path_parts) or "index"
            filename = unquote(filename).replace("?", "_").replace("&", "_") + ".md"
            filepath = os.path.join(output_folder, filename)

            with open(filepath, "w", encoding="utf-8") as f:
                f.write(result.markdown)

            print(f"Saved: {filepath} (from {result.url})")

    print(
        f"\nCrawl complete. {len(results)} pages processed and saved to '{output_folder}' folder."
    )


if __name__ == "__main__":
    asyncio.run(main())
