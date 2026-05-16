#!/usr/bin/env python3
"""Build the Apollo Windows 3-Month Marketing Roadmap PPTX."""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pptx_builder.core import *
from pptx_builder.slides import *  # noqa
from pptx_builder.slides import _textbox, _rect, _rounded_rect, wrap_slide
import zipfile

OUTPUT_FILE = "Apollo_Windows_3Month_Marketing_Roadmap.pptx"


def slide_cover():
    """Slide 1: Title/Cover slide with dark blue background."""
    shapes = ""
    # Accent bar at top
    shapes += _rect(2, "top_bar", 0, 0, SLIDE_WIDTH, inches(0.15), "E74C3C")
    # Title
    shapes += _textbox(3, "title", inches(1.5), inches(1.8), inches(7), inches(1.5), [
        {"text": "3-MONTH DIGITAL MARKETING", "size": 3200, "bold": True, "color": "FFFFFF", "align": "l"},
        {"text": "ROADMAP & STRATEGY", "size": 3200, "bold": True, "color": "E74C3C", "align": "l", "spc_before": 200},
    ])
    # Subtitle
    shapes += _textbox(4, "subtitle", inches(1.5), inches(3.8), inches(7), inches(1.2), [
        {"text": "Apollo Windows & Doors | Calgary, AB", "size": 1800, "color": "BDC3C7", "align": "l"},
        {"text": "Prepared: May 2026", "size": 1400, "color": "7F8C8D", "align": "l", "spc_before": 300},
    ])
    # Right side info box
    shapes += _rounded_rect(5, "info_box", inches(8.5), inches(2.0), inches(3.2), inches(3.5), "2C3E50", [
        {"text": "INCLUDES", "size": 1100, "bold": True, "color": "E74C3C"},
        {"text": "", "size": 600, "color": "FFFFFF"},
        {"text": "Website Audit", "size": 1000, "color": "FFFFFF"},
        {"text": "Competitor Analysis", "size": 1000, "color": "FFFFFF"},
        {"text": "Social Media Plan", "size": 1000, "color": "FFFFFF"},
        {"text": "Content Calendar", "size": 1000, "color": "FFFFFF"},
        {"text": "Ads Strategy", "size": 1000, "color": "FFFFFF"},
        {"text": "Pricing Proposal", "size": 1000, "color": "FFFFFF"},
    ])
    # Bottom bar
    shapes += _rect(6, "bottom_bar", 0, inches(7.2), SLIDE_WIDTH, inches(0.3), "E74C3C")
    return wrap_slide(shapes, "1B2A4A")


def slide_toc():
    """Slide 2: Table of Contents."""
    shapes = ""
    shapes += _rect(2, "header_bg", 0, 0, SLIDE_WIDTH, inches(1.2), "1B2A4A")
    shapes += _textbox(3, "header", inches(0.5), inches(0.2), inches(8), inches(0.9), [
        {"text": "AGENDA / TABLE OF CONTENTS", "size": 2200, "bold": True, "color": "FFFFFF", "align": "l"},
    ])
    # Left column items
    items_left = [
        "01  Brand Overview & Current State",
        "02  Website Audit & SEO Issues",
        "03  Competitor Analysis",
        "04  Social Media Audit",
        "05  Opportunities & Quick Wins",
    ]
    items_right = [
        "06  3-Month Content Plan",
        "07  Ads Strategy (Engagement & Growth)",
        "08  Month-by-Month Roadmap",
        "09  KPIs & Success Metrics",
        "10  Standard Plan Proposal",
    ]
    left_paras = [{"text": item, "size": 1400, "color": "2C3E50", "spc_before": 400} for item in items_left]
    right_paras = [{"text": item, "size": 1400, "color": "2C3E50", "spc_before": 400} for item in items_right]
    shapes += _textbox(4, "left_col", inches(0.8), inches(1.6), inches(5), inches(5), left_paras)
    shapes += _textbox(5, "right_col", inches(6.2), inches(1.6), inches(5), inches(5), right_paras)
    # Accent line
    shapes += _rect(6, "accent", inches(0.8), inches(1.4), inches(2), inches(0.04), "E74C3C")
    return wrap_slide(shapes)


def slide_brand_overview():
    """Slide 3: Brand Overview."""
    shapes = ""
    shapes += _rect(2, "header_bg", 0, 0, SLIDE_WIDTH, inches(1.2), "1B2A4A")
    shapes += _textbox(3, "header", inches(0.5), inches(0.2), inches(10), inches(0.9), [
        {"text": "01 | BRAND OVERVIEW & CURRENT STATE", "size": 2000, "bold": True, "color": "FFFFFF", "align": "l"},
    ])
    # Brand info box
    shapes += _rounded_rect(4, "brand_box", inches(0.5), inches(1.5), inches(5.5), inches(5.5), "F0F4F8", [
        {"text": "APOLLO WINDOWS & DOORS", "size": 1400, "bold": True, "color": "1B2A4A"},
        {"text": "", "size": 400, "color": "1B2A4A"},
        {"text": "Est. 1999 | Family-Owned | Calgary, AB", "size": 1100, "color": "2C3E50"},
        {"text": "", "size": 400, "color": "1B2A4A"},
        {"text": "Location: 26-3530 32 St NE, Calgary", "size": 1000, "color": "555555"},
        {"text": "Phone: (403) 568-0777", "size": 1000, "color": "555555"},
        {"text": "Hours: Mon-Fri 8-5, Sat 9-2:30", "size": 1000, "color": "555555"},
        {"text": "BBB Rating: A+ Accredited", "size": 1000, "color": "555555"},
        {"text": "", "size": 400, "color": "1B2A4A"},
        {"text": "PRODUCTS:", "size": 1000, "bold": True, "color": "1B2A4A"},
        {"text": "Vinyl & Hybrid Windows", "size": 1000, "color": "555555"},
        {"text": "Swing & Patio Doors", "size": 1000, "color": "555555"},
        {"text": "Insulated Glass Solutions", "size": 1000, "color": "555555"},
        {"text": "Custom Manufacturing", "size": 1000, "color": "555555"},
    ])
    # Strengths box
    shapes += _rounded_rect(5, "strengths", inches(6.3), inches(1.5), inches(5.3), inches(2.5), "2ECC71", [
        {"text": "STRENGTHS", "size": 1200, "bold": True, "color": "FFFFFF"},
        {"text": "25+ years in business", "size": 1000, "color": "FFFFFF"},
        {"text": "Local manufacturer (Calgary-made)", "size": 1000, "color": "FFFFFF"},
        {"text": "BBB A+ rated - high trust", "size": 1000, "color": "FFFFFF"},
        {"text": "Climate-specific expertise", "size": 1000, "color": "FFFFFF"},
        {"text": "Competitive pricing", "size": 1000, "color": "FFFFFF"},
    ])
    # Challenges box
    shapes += _rounded_rect(6, "challenges", inches(6.3), inches(4.3), inches(5.3), inches(2.7), "E74C3C", [
        {"text": "CHALLENGES IDENTIFIED", "size": 1200, "bold": True, "color": "FFFFFF"},
        {"text": "Limited social media presence", "size": 1000, "color": "FFFFFF"},
        {"text": "Website lacks modern SEO structure", "size": 1000, "color": "FFFFFF"},
        {"text": "Low online visibility vs. competitors", "size": 1000, "color": "FFFFFF"},
        {"text": "No consistent content strategy", "size": 1000, "color": "FFFFFF"},
        {"text": "Minimal Google reviews strategy", "size": 1000, "color": "FFFFFF"},
    ])
    return wrap_slide(shapes)


def slide_website_audit():
    """Slide 4: Website Audit."""
    shapes = ""
    shapes += _rect(2, "header_bg", 0, 0, SLIDE_WIDTH, inches(1.2), "1B2A4A")
    shapes += _textbox(3, "header", inches(0.5), inches(0.2), inches(10), inches(0.9), [
        {"text": "02 | WEBSITE AUDIT & SEO ISSUES", "size": 2000, "bold": True, "color": "FFFFFF", "align": "l"},
    ])
    # Issues found
    shapes += _textbox(4, "issues_title", inches(0.5), inches(1.4), inches(5.5), inches(0.5), [
        {"text": "CRITICAL ISSUES TO ADDRESS", "size": 1300, "bold": True, "color": "E74C3C"},
    ])
    issues = [
        {"text": "Missing or weak meta titles & descriptions on key pages", "size": 1050, "color": "333333", "bullet": True, "spc_before": 200},
        {"text": "No active blog / content hub for SEO ranking", "size": 1050, "color": "333333", "bullet": True, "spc_before": 200},
        {"text": "Lack of location-specific landing pages (Calgary areas)", "size": 1050, "color": "333333", "bullet": True, "spc_before": 200},
        {"text": "Page speed optimization needed (images, caching)", "size": 1050, "color": "333333", "bullet": True, "spc_before": 200},
        {"text": "Missing schema markup (LocalBusiness, Product)", "size": 1050, "color": "333333", "bullet": True, "spc_before": 200},
        {"text": "No clear CTA funnel on service pages", "size": 1050, "color": "333333", "bullet": True, "spc_before": 200},
        {"text": "Limited internal linking structure", "size": 1050, "color": "333333", "bullet": True, "spc_before": 200},
        {"text": "GMB profile needs optimization & regular posts", "size": 1050, "color": "333333", "bullet": True, "spc_before": 200},
    ]
    shapes += _textbox(5, "issues", inches(0.5), inches(1.9), inches(5.5), inches(5), issues)
    # Recommendations
    shapes += _rounded_rect(6, "reco_box", inches(6.3), inches(1.4), inches(5.3), inches(5.5), "F0F4F8", [
        {"text": "SEO PRIORITY ACTIONS", "size": 1200, "bold": True, "color": "1B2A4A"},
        {"text": "", "size": 300, "color": "1B2A4A"},
        {"text": "1. Optimize all page titles & metas", "size": 1000, "color": "333333"},
        {"text": "2. Add LocalBusiness schema markup", "size": 1000, "color": "333333"},
        {"text": "3. Create Calgary neighbourhood pages", "size": 1000, "color": "333333"},
        {"text": "4. Launch blog (2 posts/month min)", "size": 1000, "color": "333333"},
        {"text": "5. Compress images & enable caching", "size": 1000, "color": "333333"},
        {"text": "6. Fix mobile responsiveness issues", "size": 1000, "color": "333333"},
        {"text": "7. Add customer testimonials pages", "size": 1000, "color": "333333"},
        {"text": "8. Weekly GMB posts with offers", "size": 1000, "color": "333333"},
        {"text": "", "size": 300, "color": "1B2A4A"},
        {"text": "TARGET KEYWORDS:", "size": 1000, "bold": True, "color": "E74C3C"},
        {"text": "windows calgary, window replacement", "size": 900, "color": "555555"},
        {"text": "energy efficient windows alberta", "size": 900, "color": "555555"},
        {"text": "door installation calgary", "size": 900, "color": "555555"},
        {"text": "vinyl windows near me", "size": 900, "color": "555555"},
    ])
    return wrap_slide(shapes)


if __name__ == "__main__":
    from build_slides_part2 import *
    from build_slides_part3 import *

    slides = [
        slide_cover(),
        slide_toc(),
        slide_brand_overview(),
        slide_website_audit(),
        slide_competitor_analysis(),
        slide_social_audit(),
        slide_opportunities(),
        slide_content_plan_month1(),
        slide_content_plan_month2(),
        slide_content_plan_month3(),
        slide_ads_strategy(),
        slide_roadmap_overview(),
        slide_kpis(),
        slide_proposal(),
        slide_thank_you(),
    ]

    num_slides = len(slides)

    # Build content types
    overrides = "\n".join([
        f'  <Override PartName="/ppt/slides/slide{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(num_slides)
    ])
    content_types = create_content_types().replace("{slide_overrides}", overrides)

    # Write PPTX
    with zipfile.ZipFile(OUTPUT_FILE, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", create_rels())
        zf.writestr("ppt/presentation.xml", create_presentation(num_slides))
        zf.writestr("ppt/_rels/presentation.xml.rels", create_presentation_rels(num_slides))
        zf.writestr("ppt/theme/theme1.xml", create_theme())
        zf.writestr("ppt/slideMasters/slideMaster1.xml", create_slide_master())
        zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", create_slide_master_rels())
        zf.writestr("ppt/slideLayouts/slideLayout1.xml", create_slide_layout())
        zf.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", create_slide_layout_rels())

        for i, slide_xml in enumerate(slides):
            zf.writestr(f"ppt/slides/slide{i+1}.xml", slide_xml)
            zf.writestr(f"ppt/slides/_rels/slide{i+1}.xml.rels", create_slide_rels())

    print(f"SUCCESS: Created {OUTPUT_FILE} with {num_slides} slides")
