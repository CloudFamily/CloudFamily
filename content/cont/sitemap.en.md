---
date: 2018-11-29T08:41:44+01:00
title: Sitemap
weight: 40
tags: ["documentation", "tutorial"]
sitemapExclude: true
---

We should be careful about the references we include in the sitemap xml file, to avoid a high volume of pages included in the sitemap and consequently being indexed to Search Engines such as Google, Bing and more. Especially because we should mind about keeping the most relevant content in the top of the search engine index.

## Configuration 

Hugo by default indexes all pages to the sitemap. 

The proposal is just to include a parameter to the pages we want to exclude from sitemap.
So, you just add the parameter *sitemapExclude* to the given page: 

```markdown
---
date: 2018-11-29T08:41:44+01:00
title: Theme tutorial
weight: 15
tags: ["tutorial", "theme"] 
sitemapExclude: true
---
```