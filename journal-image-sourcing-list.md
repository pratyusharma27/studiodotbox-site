# Journal image sourcing list

All URLs are Unsplash (free license, no attribution required), pre-sized to 1800×1200 via Unsplash's own CDN params. Same image is used for hero and thumb per your spec — just save the URL twice under the two filenames.

Download each URL directly (right-click → Save As, or `curl -o filename url`) into `images/journal/`.

| Article | Theme matched | Image URL (1800×1200 .jpg) | Save as |
|---|---|---|---|
| architecture-and-wellbeing | Interior atrium, tree, natural light | https://images.unsplash.com/photo-1760225137568-256ae231ccbe?w=1800&h=1200&fit=crop&q=80 | architecture-and-wellbeing.jpg + architecture-and-wellbeing-thumb.jpg |
| designing-for-climate-context | Aerial view, building with green roof | https://images.unsplash.com/photo-1642833714391-cf9a0732321d?w=1800&h=1200&fit=crop&q=80 | designing-for-climate-context.jpg + -thumb.jpg |
| designing-for-india | Lotus Temple, modern Indian architecture | https://images.unsplash.com/photo-1761480458272-4b2f06cda520?w=1800&h=1200&fit=crop&q=80 | designing-for-india.jpg + -thumb.jpg |
| geometry-and-meaning | B&W concrete structure, strong geometry | https://images.unsplash.com/photo-1559763194-521eef49b386?w=1800&h=1200&fit=crop&q=80 | geometry-and-meaning.jpg + -thumb.jpg |
| hiring-architecture-talent | Team meeting, modern conference room | https://images.unsplash.com/photo-1769739576456-0aefcff3f4b9?w=1800&h=1200&fit=crop&q=80 | hiring-architecture-talent.jpg + -thumb.jpg |
| living-with-imperfection | Weathered concrete wall, cracks | https://images.unsplash.com/photo-1635284305301-e485d088bf94?w=1800&h=1200&fit=crop&q=80 | living-with-imperfection.jpg + -thumb.jpg |
| personal-brand-as-architect | Architect drafting with pencil and ruler | https://images.unsplash.com/photo-1503387762-592deb58ef4e?w=1800&h=1200&fit=crop&q=80 | personal-brand-as-architect.jpg + -thumb.jpg |
| returning-to-craft | Hand-cutting wood, craftsmanship | https://images.unsplash.com/photo-1659930087003-2d64e33181f7?w=1800&h=1200&fit=crop&q=80 | returning-to-craft.jpg + -thumb.jpg |
| scaling-design-without-losing-craft | Retail store interior | https://images.unsplash.com/photo-1441986300917-64674bd600d8?w=1800&h=1200&fit=crop&q=80 | scaling-design-without-losing-craft.jpg + -thumb.jpg |
| site-analysis-or-site-worship | Green lawn, bushes, site landscape | https://images.unsplash.com/photo-1668120089662-42642838cfef?w=1800&h=1200&fit=crop&q=80 | site-analysis-or-site-worship.jpg + -thumb.jpg |
| sustainability-beyond-buzzwords | Concrete building covered in trees | https://images.unsplash.com/photo-1451976426598-a7593bd6d0b2?w=1800&h=1200&fit=crop&q=80 | sustainability-beyond-buzzwords.jpg + -thumb.jpg |
| the-aerotropolis-dream | Aerial view, airport with airplanes | https://images.unsplash.com/photo-1466691623998-d607fab1ca29?w=1800&h=1200&fit=crop&q=80 | the-aerotropolis-dream.jpg + -thumb.jpg |
| the-brief-is-never-the-actual-problem | People reviewing architectural blueprints | https://images.unsplash.com/photo-1608303588026-884930af2559?w=1800&h=1200&fit=crop&q=80 | the-brief-is-never-the-actual-problem.jpg + -thumb.jpg |
| the-collaboration-gap | Crane near building, construction site | https://images.unsplash.com/photo-1599707254554-027aeb4deacd?w=1800&h=1200&fit=crop&q=80 | the-collaboration-gap.jpg + -thumb.jpg |
| the-material-story | B&W marble surface texture | https://images.unsplash.com/photo-1566305977571-5666677c6e98?w=1800&h=1200&fit=crop&q=80 | the-material-story.jpg + -thumb.jpg |
| the-undesigned-moments | Shadow on side of a building | https://images.unsplash.com/photo-1563891925196-a3e34f6d869c?w=1800&h=1200&fit=crop&q=80 | the-undesigned-moments.jpg + -thumb.jpg |
| the-unfinished-nature-of-practice | Person drafting on blueprint | https://images.unsplash.com/photo-1503387837-b154d5074bd2?w=1800&h=1200&fit=crop&q=80 | the-unfinished-nature-of-practice.jpg + -thumb.jpg |
| the-view-problem | Room with wood paneling, large windows | https://images.unsplash.com/photo-1765371512707-9e0e96fd9e5b?w=1800&h=1200&fit=crop&q=80 | the-view-problem.jpg + -thumb.jpg |
| when-budget-becomes-creativity | Gray minimalist building | https://images.unsplash.com/photo-1496236436299-cde70e3587cf?w=1800&h=1200&fit=crop&q=80 | when-budget-becomes-creativity.jpg + -thumb.jpg |
| when-clients-teach-you | Man and woman in discussion at table | https://images.unsplash.com/photo-1616587656879-8a8a7cce9d6c?w=1800&h=1200&fit=crop&q=80 | when-clients-teach-you.jpg + -thumb.jpg |
| working-across-scales | Architectural model of a city plan | https://images.unsplash.com/photo-1741163269578-c51176eddb3c?w=1800&h=1200&fit=crop&q=80 | working-across-scales.jpg + -thumb.jpg |

## Notes

- All source images are free-tier Unsplash (not Unsplash+), so no paywall on download.
- The `w=1800&h=1200&fit=crop` params are Unsplash's own image API — the URL itself returns an already-cropped 1800×1200 JPG, no editing needed on your end.
- Each article gets two identical files (hero + thumb) per your original spec.
- If any crop looks off on a specific photo (Unsplash's auto-crop centers by default), add `&crop=entropy` to the URL for smarter subject-aware cropping.
