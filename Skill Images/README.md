# Skill Images

Every image in this folder becomes a floating skill orb on the Scroll Story
opening screen. The orbs swap to a different skill every 2–5 seconds.

**The file name is the skill name.** Name each file exactly as the skill is
written in the `Skills` column of `Employee Details.xlsx` (case doesn't matter),
e.g. `Power BI.svg`, `SQL.png`. When it matches, the orb shows how many people
hold that skill. When it doesn't (e.g. `AWS` today), the orb shows only the name.

Formats: .svg (best), .png, .jpg, .webp. Square images with transparent
backgrounds look best, because the orbs are dark circles.

To add a skill: drop a file in. To remove one: delete it. Then rebuild:

    cd "../../02_Scripts & ETL"
    python build_scroll_story.py

## What's here now

| File | Source |
|---|---|
| Python.svg, Azure.svg, SQL.svg (SQL Server mark), AWS.svg | Devicon 2.17, MIT licence. AWS text recoloured white for the dark background |
| Databricks.svg, Snowflake.svg | Simple Icons 16.31, CC0, in each brand's own colour |
| Alteryx.svg, Power BI.svg, Excel.svg, Tableau.svg | **Placeholders** (the name as text). These logos aren't in open icon sets. Replace each with the official logo from the vendor's brand/press page or PwC's approved asset library, keeping the file name |

Tool logos are the property of their owners. They are used here only to label skills on an internal page.
