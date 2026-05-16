#!/usr/bin/env python3
"""Build Apollo Windows 3-Month Marketing Roadmap PPTX - Brand Colors Edition."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx_builder.core import inches, SLIDE_WIDTH, SLIDE_HEIGHT, build_pptx
from pptx_builder.slides import (
    _textbox, _rect, _rounded_rect, wrap_slide,
    RED, NAVY, DARK, WHITE, LIGHT_BG, LIGHT_RED, GREY, TEXT, SUBTEXT
)

OUTPUT = "Apollo_Windows_3Month_Marketing_Roadmap.pptx"


def slide_cover():
    s = ""
    s += _rect(2, "top", 0, 0, SLIDE_WIDTH, inches(0.12), RED)
    s += _textbox(3, "t", inches(1.2), inches(1.6), inches(7.5), inches(1.8), [
        {"text": "3-MONTH DIGITAL MARKETING", "size": 3400, "bold": True, "color": WHITE, "align": "l"},
        {"text": "ROADMAP & STRATEGY", "size": 3400, "bold": True, "color": RED, "align": "l", "spc_before": 200},
    ])
    s += _textbox(4, "sub", inches(1.2), inches(3.8), inches(7), inches(1.2), [
        {"text": "Apollo Windows & Doors | Calgary, AB", "size": 1800, "color": "BDC3C7", "align": "l"},
        {"text": "Prepared: May 2026", "size": 1300, "color": GREY, "align": "l", "spc_before": 300},
    ])
    s += _rounded_rect(5, "box", inches(8.8), inches(1.8), inches(3.0), inches(3.8), DARK, [
        {"text": "INCLUDES", "size": 1100, "bold": True, "color": RED},
        {"text": "", "size": 500, "color": WHITE},
        {"text": "Website Audit", "size": 1000, "color": WHITE},
        {"text": "Competitor Analysis", "size": 1000, "color": WHITE},
        {"text": "Social Media Plan", "size": 1000, "color": WHITE},
        {"text": "Content Calendar", "size": 1000, "color": WHITE},
        {"text": "Ads Strategy", "size": 1000, "color": WHITE},
        {"text": "Service Proposal", "size": 1000, "color": WHITE},
    ])
    s += _rect(6, "bot", 0, inches(7.3), SLIDE_WIDTH, inches(0.2), RED)
    return wrap_slide(s, NAVY)


def slide_toc():
    s = ""
    s += _rect(2, "h", 0, 0, SLIDE_WIDTH, inches(1.1), NAVY)
    s += _textbox(3, "ht", inches(0.5), inches(0.15), inches(9), inches(0.9), [
        {"text": "TABLE OF CONTENTS", "size": 2200, "bold": True, "color": WHITE},
    ])
    left = ["01  Brand Overview & Current State", "02  Website Audit & SEO Issues",
            "03  Competitor Analysis", "04  Social Media Audit",
            "05  Opportunities & Quick Wins", "06  Content Plan (Month 1-3)"]
    right = ["07  Ads Strategy ($300/mo Breakdown)",
             "08  Month-by-Month Roadmap", "09  KPIs & Success Metrics",
             "10  Standard Plan (INR 10K)", "11  Business Plan (INR 18K)"]
    lp = [{"text": i, "size": 1350, "color": DARK, "spc_before": 350} for i in left]
    rp = [{"text": i, "size": 1350, "color": DARK, "spc_before": 350} for i in right]
    s += _textbox(4, "l", inches(0.7), inches(1.4), inches(5.5), inches(5.5), lp)
    s += _textbox(5, "r", inches(6.3), inches(1.4), inches(5.5), inches(5.5), rp)
    s += _rect(6, "acc", inches(0.7), inches(1.25), inches(1.8), inches(0.04), RED)
    return wrap_slide(s)


def slide_brand_overview():
    s = ""
    s += _rect(2, "h", 0, 0, SLIDE_WIDTH, inches(1.1), NAVY)
    s += _textbox(3, "ht", inches(0.5), inches(0.15), inches(10), inches(0.9), [
        {"text": "01 | BRAND OVERVIEW & CURRENT STATE", "size": 1900, "bold": True, "color": WHITE},
    ])
    s += _rounded_rect(4, "info", inches(0.4), inches(1.4), inches(5.6), inches(5.7), LIGHT_BG, [
        {"text": "APOLLO WINDOWS & DOORS", "size": 1400, "bold": True, "color": NAVY},
        {"text": "", "size": 350, "color": NAVY},
        {"text": "Est. 1999 | Family-Owned | Calgary, AB", "size": 1050, "color": DARK},
        {"text": "", "size": 350, "color": NAVY},
        {"text": "Location: 26-3530 32 St NE, Calgary", "size": 1000, "color": SUBTEXT},
        {"text": "Phone: (403) 568-0777", "size": 1000, "color": SUBTEXT},
        {"text": "Hours: Mon-Fri 8-5, Sat 9-2:30", "size": 1000, "color": SUBTEXT},
        {"text": "BBB Rating: A+ Accredited", "size": 1000, "color": SUBTEXT},
        {"text": "", "size": 350, "color": NAVY},
        {"text": "PRODUCTS:", "size": 1000, "bold": True, "color": NAVY},
        {"text": "Vinyl & Hybrid Windows", "size": 1000, "color": SUBTEXT},
        {"text": "Swing & Patio Doors", "size": 1000, "color": SUBTEXT},
        {"text": "Insulated Glass Solutions", "size": 1000, "color": SUBTEXT},
        {"text": "Custom Manufacturing", "size": 1000, "color": SUBTEXT},
    ])
    s += _rounded_rect(5, "str", inches(6.2), inches(1.4), inches(5.4), inches(2.6), "1D6B3F", [
        {"text": "STRENGTHS", "size": 1200, "bold": True, "color": WHITE},
        {"text": "25+ years in business", "size": 1000, "color": WHITE},
        {"text": "Local manufacturer (Calgary-made)", "size": 1000, "color": WHITE},
        {"text": "BBB A+ rated - high trust signal", "size": 1000, "color": WHITE},
        {"text": "Climate-specific engineering", "size": 1000, "color": WHITE},
        {"text": "Competitive pricing vs. national brands", "size": 1000, "color": WHITE},
    ])
    s += _rounded_rect(6, "ch", inches(6.2), inches(4.2), inches(5.4), inches(2.9), RED, [
        {"text": "CHALLENGES IDENTIFIED", "size": 1200, "bold": True, "color": WHITE},
        {"text": "Limited social media presence", "size": 1000, "color": WHITE},
        {"text": "Website lacks modern SEO structure", "size": 1000, "color": WHITE},
        {"text": "Low online visibility vs. competitors", "size": 1000, "color": WHITE},
        {"text": "No consistent content strategy", "size": 1000, "color": WHITE},
        {"text": "Minimal Google reviews strategy", "size": 1000, "color": WHITE},
        {"text": "No paid ads running currently", "size": 1000, "color": WHITE},
    ])
    return wrap_slide(s)


def slide_website_audit():
    s = ""
    s += _rect(2, "h", 0, 0, SLIDE_WIDTH, inches(1.1), NAVY)
    s += _textbox(3, "ht", inches(0.5), inches(0.15), inches(10), inches(0.9), [
        {"text": "02 | WEBSITE AUDIT & SEO ISSUES", "size": 1900, "bold": True, "color": WHITE},
    ])
    s += _textbox(4, "it", inches(0.5), inches(1.3), inches(5.5), inches(0.4), [
        {"text": "CRITICAL ISSUES TO ADDRESS", "size": 1250, "bold": True, "color": RED},
    ])
    issues = [
        "Missing/weak meta titles & descriptions",
        "No active blog or content hub for SEO",
        "No location-specific landing pages",
        "Page speed optimization needed",
        "Missing schema markup (LocalBusiness)",
        "No clear CTA funnel on service pages",
        "Limited internal linking structure",
        "GMB profile needs optimization & posts",
    ]
    ip = [{"text": t, "size": 1000, "color": TEXT, "bullet": True, "spc_before": 180} for t in issues]
    s += _textbox(5, "iss", inches(0.5), inches(1.8), inches(5.5), inches(5.2), ip)
    s += _rounded_rect(6, "reco", inches(6.2), inches(1.3), inches(5.4), inches(5.8), LIGHT_BG, [
        {"text": "SEO PRIORITY ACTIONS", "size": 1200, "bold": True, "color": NAVY},
        {"text": "", "size": 250, "color": NAVY},
        {"text": "1. Optimize all page titles & metas", "size": 1000, "color": TEXT},
        {"text": "2. Add LocalBusiness schema markup", "size": 1000, "color": TEXT},
        {"text": "3. Create Calgary neighbourhood pages", "size": 1000, "color": TEXT},
        {"text": "4. Launch blog (2 posts/month)", "size": 1000, "color": TEXT},
        {"text": "5. Compress images & enable caching", "size": 1000, "color": TEXT},
        {"text": "6. Add testimonials page", "size": 1000, "color": TEXT},
        {"text": "7. Weekly GMB posts with offers", "size": 1000, "color": TEXT},
        {"text": "", "size": 300, "color": NAVY},
        {"text": "TARGET KEYWORDS:", "size": 1000, "bold": True, "color": RED},
        {"text": "windows calgary", "size": 950, "color": SUBTEXT},
        {"text": "window replacement calgary", "size": 950, "color": SUBTEXT},
        {"text": "energy efficient windows alberta", "size": 950, "color": SUBTEXT},
        {"text": "door installation calgary", "size": 950, "color": SUBTEXT},
        {"text": "vinyl windows near me", "size": 950, "color": SUBTEXT},
    ])
    return wrap_slide(s)


def slide_competitors():
    s = ""
    s += _rect(2, "h", 0, 0, SLIDE_WIDTH, inches(1.1), NAVY)
    s += _textbox(3, "ht", inches(0.5), inches(0.15), inches(10), inches(0.9), [
        {"text": "03 | COMPETITOR ANALYSIS", "size": 1900, "bold": True, "color": WHITE},
    ])
    comps = [
        ("Lux Windows", RED, "Premium positioning, active social, video content, showroom experience"),
        ("All Weather", "1D6B3F", "10x Energy Star winner, strong blog/SEO, dealer network, Calgary+Edmonton"),
        ("Durabuilt", "D4740E", "Multi-city, design galleries, social channels, builder partnerships"),
    ]
    y = inches(1.4)
    for nm, cl, desc in comps:
        s += _rounded_rect(10, f"c_{nm}", inches(0.4), y, inches(3.6), inches(1.6), cl, [
            {"text": nm, "size": 1200, "bold": True, "color": WHITE},
            {"text": desc, "size": 850, "color": WHITE},
        ])
        y += inches(1.8)
    comps2 = [
        ("JELD-WEN", "5B2D8E", "National brand, massive catalog, high domain authority, content machine"),
        ("WindowMart", "2980B9", "Aggressive pricing, strong Google Ads, review generation, multiple locations"),
        ("Windows Canada", "1ABC9C", "Modern site, project galleries, financing, active blog, local SEO"),
    ]
    y = inches(1.4)
    for nm, cl, desc in comps2:
        s += _rounded_rect(11, f"c2_{nm}", inches(4.2), y, inches(3.6), inches(1.6), cl, [
            {"text": nm, "size": 1200, "bold": True, "color": WHITE},
            {"text": desc, "size": 850, "color": WHITE},
        ])
        y += inches(1.8)
    s += _rounded_rect(20, "tk", inches(8.0), inches(1.4), inches(3.7), inches(5.6), NAVY, [
        {"text": "KEY GAPS", "size": 1100, "bold": True, "color": RED},
        {"text": "", "size": 250, "color": WHITE},
        {"text": "Apollo lacks:", "size": 1000, "color": WHITE},
        {"text": "Active blog content", "size": 900, "color": "BDC3C7"},
        {"text": "Video/Reels presence", "size": 900, "color": "BDC3C7"},
        {"text": "Project galleries", "size": 900, "color": "BDC3C7"},
        {"text": "Review generation system", "size": 900, "color": "BDC3C7"},
        {"text": "Social engagement", "size": 900, "color": "BDC3C7"},
        {"text": "Paid ads strategy", "size": 900, "color": "BDC3C7"},
        {"text": "", "size": 300, "color": WHITE},
        {"text": "OPPORTUNITY:", "size": 1000, "bold": True, "color": "2ECC71"},
        {"text": "Local trust + family", "size": 900, "color": WHITE},
        {"text": "story = differentiator", "size": 900, "color": WHITE},
    ])
    return wrap_slide(s)


def slide_social_audit():
    s = ""
    s += _rect(2, "h", 0, 0, SLIDE_WIDTH, inches(1.1), NAVY)
    s += _textbox(3, "ht", inches(0.5), inches(0.15), inches(10), inches(0.9), [
        {"text": "04 | SOCIAL MEDIA AUDIT", "size": 1900, "bold": True, "color": WHITE},
    ])
    s += _rounded_rect(4, "ig", inches(0.4), inches(1.4), inches(5.6), inches(5.7), LIGHT_RED, [
        {"text": "INSTAGRAM @apollowindowscalgary", "size": 1150, "bold": True, "color": RED},
        {"text": "", "size": 250, "color": TEXT},
        {"text": "Current Status:", "size": 1050, "bold": True, "color": TEXT},
        {"text": "Low follower count & engagement", "size": 1000, "color": SUBTEXT},
        {"text": "Inconsistent posting frequency", "size": 1000, "color": SUBTEXT},
        {"text": "Limited use of Reels/Stories", "size": 1000, "color": SUBTEXT},
        {"text": "No hashtag strategy", "size": 1000, "color": SUBTEXT},
        {"text": "Missing bio link optimization", "size": 1000, "color": SUBTEXT},
        {"text": "", "size": 250, "color": TEXT},
        {"text": "Issues:", "size": 1050, "bold": True, "color": RED},
        {"text": "No content pillars defined", "size": 1000, "color": SUBTEXT},
        {"text": "No community engagement loop", "size": 1000, "color": SUBTEXT},
        {"text": "Product-only posts (no storytelling)", "size": 1000, "color": SUBTEXT},
        {"text": "No user-generated content", "size": 1000, "color": SUBTEXT},
    ])
    s += _rounded_rect(5, "fb", inches(6.2), inches(1.4), inches(5.4), inches(5.7), "E8F4FD", [
        {"text": "FACEBOOK @apollowindowscalgary", "size": 1150, "bold": True, "color": "2980B9"},
        {"text": "", "size": 250, "color": TEXT},
        {"text": "Current Status:", "size": 1050, "bold": True, "color": TEXT},
        {"text": "Page exists but underutilized", "size": 1000, "color": SUBTEXT},
        {"text": "Low engagement rate on posts", "size": 1000, "color": SUBTEXT},
        {"text": "Few customer reviews on page", "size": 1000, "color": SUBTEXT},
        {"text": "No Facebook Events or Offers used", "size": 1000, "color": SUBTEXT},
        {"text": "Limited community interaction", "size": 1000, "color": SUBTEXT},
        {"text": "", "size": 250, "color": TEXT},
        {"text": "Issues:", "size": 1050, "bold": True, "color": RED},
        {"text": "No paid promotion strategy", "size": 1000, "color": SUBTEXT},
        {"text": "Missing Facebook Shop/catalog", "size": 1000, "color": SUBTEXT},
        {"text": "No video content or Lives", "size": 1000, "color": SUBTEXT},
        {"text": "Inconsistent brand voice", "size": 1000, "color": SUBTEXT},
    ])
    return wrap_slide(s)


def slide_opportunities():
    s = ""
    s += _rect(2, "h", 0, 0, SLIDE_WIDTH, inches(1.1), NAVY)
    s += _textbox(3, "ht", inches(0.5), inches(0.15), inches(10), inches(0.9), [
        {"text": "05 | OPPORTUNITIES & QUICK WINS", "size": 1900, "bold": True, "color": WHITE},
    ])
    s += _rounded_rect(4, "qw", inches(0.4), inches(1.4), inches(5.6), inches(2.9), "1D6B3F", [
        {"text": "QUICK WINS (Week 1-2)", "size": 1200, "bold": True, "color": WHITE},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Optimize GMB profile completely", "size": 1000, "color": WHITE},
        {"text": "Fix website meta titles & descriptions", "size": 1000, "color": WHITE},
        {"text": "Set up consistent posting schedule", "size": 1000, "color": WHITE},
        {"text": "Create branded hashtag set", "size": 1000, "color": WHITE},
        {"text": "Ask 10 past clients for Google reviews", "size": 1000, "color": WHITE},
        {"text": "Update Instagram bio + link tree", "size": 1000, "color": WHITE},
    ])
    s += _rounded_rect(5, "sm", inches(0.4), inches(4.5), inches(5.6), inches(2.6), "2980B9", [
        {"text": "SOCIAL MEDIA OPPORTUNITIES", "size": 1200, "bold": True, "color": WHITE},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Before/After project reveals (high engagement)", "size": 1000, "color": WHITE},
        {"text": "Behind-the-scenes manufacturing Reels", "size": 1000, "color": WHITE},
        {"text": "Calgary homeowner tips (seasonal)", "size": 1000, "color": WHITE},
        {"text": "Customer testimonial videos", "size": 1000, "color": WHITE},
        {"text": "Local partnership cross-promos", "size": 1000, "color": WHITE},
    ])
    s += _rounded_rect(6, "gr", inches(6.2), inches(1.4), inches(5.4), inches(5.7), "D4740E", [
        {"text": "GROWTH OPPORTUNITIES", "size": 1200, "bold": True, "color": WHITE},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "LOCAL SEO:", "size": 1050, "bold": True, "color": WHITE},
        {"text": "Rank for 'windows Calgary' keywords", "size": 1000, "color": WHITE},
        {"text": "Target neighbourhood-specific searches", "size": 1000, "color": WHITE},
        {"text": "Build local citation consistency", "size": 1000, "color": WHITE},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "CONTENT:", "size": 1050, "bold": True, "color": WHITE},
        {"text": "Energy efficiency educational content", "size": 1000, "color": WHITE},
        {"text": "Seasonal maintenance guides", "size": 1000, "color": WHITE},
        {"text": "Product comparison guides", "size": 1000, "color": WHITE},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "PARTNERSHIPS:", "size": 1050, "bold": True, "color": WHITE},
        {"text": "Local builders & contractors", "size": 1000, "color": WHITE},
        {"text": "Home renovation influencers", "size": 1000, "color": WHITE},
        {"text": "Calgary real estate agents", "size": 1000, "color": WHITE},
    ])
    return wrap_slide(s)


if __name__ == "__main__":
    from build_slides_part2 import (
        slide_content_m1, slide_content_m2, slide_content_m3
    )
    from build_slides_part3 import (
        slide_ads_strategy, slide_roadmap, slide_kpis,
        slide_proposal_standard, slide_proposal_business, slide_thank_you
    )

    slides = [
        slide_cover(),
        slide_toc(),
        slide_brand_overview(),
        slide_website_audit(),
        slide_competitors(),
        slide_social_audit(),
        slide_opportunities(),
        slide_content_m1(),
        slide_content_m2(),
        slide_content_m3(),
        slide_ads_strategy(),
        slide_roadmap(),
        slide_kpis(),
        slide_proposal_standard(),
        slide_proposal_business(),
        slide_thank_you(),
    ]
    build_pptx(slides, OUTPUT)
