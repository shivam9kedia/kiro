"""Ads Strategy, Roadmap, KPIs, Proposals, Thank You slides."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx_builder.core import inches, SLIDE_WIDTH, SLIDE_HEIGHT
from pptx_builder.slides import (
    _textbox, _rect, _rounded_rect, wrap_slide,
    RED, NAVY, DARK, WHITE, LIGHT_BG, GREY, TEXT, SUBTEXT
)


def slide_ads_strategy():
    s = ""
    s += _rect(2, "h", 0, 0, SLIDE_WIDTH, inches(1.1), NAVY)
    s += _textbox(3, "ht", inches(0.5), inches(0.15), inches(10), inches(0.9), [
        {"text": "07 | ADS STRATEGY - $300 USD/MONTH BREAKDOWN", "size": 1900, "bold": True, "color": WHITE},
    ])
    # Profile Visiting Ads - $200
    s += _rounded_rect(4, "pv", inches(0.4), inches(1.4), inches(5.6), inches(3.2), "2980B9", [
        {"text": "PROFILE VISITING ADS - $200/month", "size": 1200, "bold": True, "color": WHITE},
        {"text": "", "size": 200, "color": WHITE},
        {"text": "Objective: Profile Visits / Audience Building", "size": 1000, "color": WHITE},
        {"text": "Goal: Build relevant local audience base", "size": 1000, "color": WHITE},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Targeting:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "Calgary homeowners, age 30-65", "size": 950, "color": "D6EAF8"},
        {"text": "Interests: Home improvement, renovation", "size": 950, "color": "D6EAF8"},
        {"text": "Radius: Calgary + 50km surrounding", "size": 950, "color": "D6EAF8"},
        {"text": "Lookalike: Similar to existing customers", "size": 950, "color": "D6EAF8"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Realistic Goal: +150-250 new followers/month", "size": 1000, "bold": True, "color": WHITE},
    ])
    # Post Engagement Ads - $100
    s += _rounded_rect(5, "pe", inches(0.4), inches(4.8), inches(5.6), inches(2.3), "1D6B3F", [
        {"text": "POST ENGAGEMENT ADS - $100/month", "size": 1200, "bold": True, "color": WHITE},
        {"text": "", "size": 200, "color": WHITE},
        {"text": "Objective: Post Engagement", "size": 1000, "color": WHITE},
        {"text": "Goal: Let relevant audience connect with content", "size": 1000, "color": WHITE},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Best Content to Boost:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "Before/After carousels, Reels, Testimonials", "size": 950, "color": "D5F5E3"},
        {"text": "Boost top 1-2 performing posts weekly", "size": 950, "color": "D5F5E3"},
    ])
    # Monthly breakdown
    s += _rounded_rect(6, "cal", inches(6.2), inches(1.4), inches(5.4), inches(5.7), LIGHT_BG, [
        {"text": "3-MONTH AD CALENDAR", "size": 1200, "bold": True, "color": NAVY},
        {"text": "", "size": 250, "color": TEXT},
        {"text": "MONTH 1 ($300):", "size": 1050, "bold": True, "color": RED},
        {"text": "$200 - Profile visits (audience building)", "size": 950, "color": SUBTEXT},
        {"text": "$100 - Boost brand story posts", "size": 950, "color": SUBTEXT},
        {"text": "Focus: Awareness + first followers", "size": 950, "color": SUBTEXT},
        {"text": "", "size": 200, "color": TEXT},
        {"text": "MONTH 2 ($300):", "size": 1050, "bold": True, "color": "D4740E"},
        {"text": "$200 - Profile visits (refine targeting)", "size": 950, "color": SUBTEXT},
        {"text": "$100 - Boost Reels + testimonials", "size": 950, "color": SUBTEXT},
        {"text": "Focus: Engagement + video views", "size": 950, "color": SUBTEXT},
        {"text": "", "size": 200, "color": TEXT},
        {"text": "MONTH 3 ($300):", "size": 1050, "bold": True, "color": "1D6B3F"},
        {"text": "$200 - Profile visits (lookalike audience)", "size": 950, "color": SUBTEXT},
        {"text": "$100 - Boost offer/CTA posts", "size": 950, "color": SUBTEXT},
        {"text": "Focus: Warm audience + conversions", "size": 950, "color": SUBTEXT},
        {"text": "", "size": 250, "color": TEXT},
        {"text": "TOTAL 3-MONTH AD SPEND: $900 USD", "size": 1100, "bold": True, "color": NAVY},
    ])
    return wrap_slide(s)


def slide_roadmap():
    s = ""
    s += _rect(2, "h", 0, 0, SLIDE_WIDTH, inches(1.1), NAVY)
    s += _textbox(3, "ht", inches(0.5), inches(0.15), inches(10), inches(0.9), [
        {"text": "08 | MONTH-BY-MONTH ROADMAP", "size": 1900, "bold": True, "color": WHITE},
    ])
    s += _rounded_rect(4, "m1", inches(0.3), inches(1.4), inches(3.7), inches(5.7), RED, [
        {"text": "MONTH 1", "size": 1400, "bold": True, "color": WHITE},
        {"text": "FOUNDATION", "size": 1100, "color": "FFDDD5"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Website:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "Meta tags optimization", "size": 900, "color": "FFDDD5"},
        {"text": "Schema markup added", "size": 900, "color": "FFDDD5"},
        {"text": "Speed optimization", "size": 900, "color": "FFDDD5"},
        {"text": "", "size": 100, "color": WHITE},
        {"text": "Social:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "Profile optimization", "size": 900, "color": "FFDDD5"},
        {"text": "Content pillars set", "size": 900, "color": "FFDDD5"},
        {"text": "3-4 posts/week start", "size": 900, "color": "FFDDD5"},
        {"text": "", "size": 100, "color": WHITE},
        {"text": "GMB:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "Full optimization", "size": 900, "color": "FFDDD5"},
        {"text": "Weekly posts begin", "size": 900, "color": "FFDDD5"},
        {"text": "Review campaign start", "size": 900, "color": "FFDDD5"},
    ])
    s += _rounded_rect(5, "m2", inches(4.15), inches(1.4), inches(3.7), inches(5.7), "D4740E", [
        {"text": "MONTH 2", "size": 1400, "bold": True, "color": WHITE},
        {"text": "ENGAGEMENT", "size": 1100, "color": "FEF3E2"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Website:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "Blog launches (2 posts)", "size": 900, "color": "FEF3E2"},
        {"text": "Location pages created", "size": 900, "color": "FEF3E2"},
        {"text": "Testimonials page", "size": 900, "color": "FEF3E2"},
        {"text": "", "size": 100, "color": WHITE},
        {"text": "Social:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "4-5 posts/week", "size": 900, "color": "FEF3E2"},
        {"text": "Reels strategy active", "size": 900, "color": "FEF3E2"},
        {"text": "Engagement campaigns", "size": 900, "color": "FEF3E2"},
        {"text": "", "size": 100, "color": WHITE},
        {"text": "Ads:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "$200 profile visits", "size": 900, "color": "FEF3E2"},
        {"text": "$100 post engagement", "size": 900, "color": "FEF3E2"},
        {"text": "A/B testing creatives", "size": 900, "color": "FEF3E2"},
    ])
    s += _rounded_rect(6, "m3", inches(8.0), inches(1.4), inches(3.7), inches(5.7), "1D6B3F", [
        {"text": "MONTH 3", "size": 1400, "bold": True, "color": WHITE},
        {"text": "CONVERSION", "size": 1100, "color": "D5F5E3"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Website:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "2 more blog posts", "size": 900, "color": "D5F5E3"},
        {"text": "CTA optimization", "size": 900, "color": "D5F5E3"},
        {"text": "Lead forms improved", "size": 900, "color": "D5F5E3"},
        {"text": "", "size": 100, "color": WHITE},
        {"text": "Social:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "Lead gen content", "size": 900, "color": "D5F5E3"},
        {"text": "Customer stories", "size": 900, "color": "D5F5E3"},
        {"text": "Conversion CTAs", "size": 900, "color": "D5F5E3"},
        {"text": "", "size": 100, "color": WHITE},
        {"text": "Ads:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "$200 lookalike targeting", "size": 900, "color": "D5F5E3"},
        {"text": "$100 boost offer posts", "size": 900, "color": "D5F5E3"},
        {"text": "Retargeting warm audience", "size": 900, "color": "D5F5E3"},
    ])
    return wrap_slide(s)


def slide_kpis():
    s = ""
    s += _rect(2, "h", 0, 0, SLIDE_WIDTH, inches(1.1), NAVY)
    s += _textbox(3, "ht", inches(0.5), inches(0.15), inches(10), inches(0.9), [
        {"text": "09 | KPIs & REALISTIC SUCCESS METRICS", "size": 1900, "bold": True, "color": WHITE},
    ])
    s += _textbox(4, "n", inches(0.5), inches(1.15), inches(11), inches(0.4), [
        {"text": "No fake promises. Achievable targets based on industry benchmarks for local service businesses.", "size": 1050, "color": GREY},
    ])
    s += _rounded_rect(5, "sk", inches(0.4), inches(1.7), inches(3.6), inches(5.3), "2980B9", [
        {"text": "SOCIAL MEDIA", "size": 1200, "bold": True, "color": WHITE},
        {"text": "", "size": 250, "color": WHITE},
        {"text": "Instagram Followers:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "+300-500 in 3 months", "size": 1000, "color": "D6EAF8"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Engagement Rate:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "Target 3-5%", "size": 1000, "color": "D6EAF8"},
        {"text": "(industry avg 1-2%)", "size": 900, "color": "D6EAF8"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Facebook Page Likes:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "+200-350 in 3 months", "size": 1000, "color": "D6EAF8"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Post Reach:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "2x-3x current reach", "size": 1000, "color": "D6EAF8"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Reels Views:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "500-2,000 per Reel avg", "size": 1000, "color": "D6EAF8"},
    ])
    s += _rounded_rect(6, "wk", inches(4.15), inches(1.7), inches(3.6), inches(5.3), "1D6B3F", [
        {"text": "WEBSITE & SEO", "size": 1200, "bold": True, "color": WHITE},
        {"text": "", "size": 250, "color": WHITE},
        {"text": "Organic Traffic:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "+30-50% increase", "size": 1000, "color": "D5F5E3"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Keyword Rankings:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "Top 20 for 5+ local terms", "size": 1000, "color": "D5F5E3"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "GMB Views:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "+40-60% increase", "size": 1000, "color": "D5F5E3"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Google Reviews:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "+10-15 new reviews", "size": 1000, "color": "D5F5E3"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Website Leads:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "+20-40% more inquiries", "size": 1000, "color": "D5F5E3"},
    ])
    s += _rounded_rect(7, "bk", inches(7.9), inches(1.7), inches(3.7), inches(5.3), "D4740E", [
        {"text": "BUSINESS IMPACT", "size": 1200, "bold": True, "color": WHITE},
        {"text": "", "size": 250, "color": WHITE},
        {"text": "Quote Requests:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "+15-25% from digital", "size": 1000, "color": "FEF3E2"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Brand Awareness:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "Measurable via reach", "size": 1000, "color": "FEF3E2"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Cost Per Follower:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "$0.80-$1.30 (with $200 budget)", "size": 1000, "color": "FEF3E2"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Cost Per Engagement:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "$0.05-$0.15 per interaction", "size": 1000, "color": "FEF3E2"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Review Rating:", "size": 1000, "bold": True, "color": WHITE},
        {"text": "Maintain 4.5+ stars", "size": 1000, "color": "FEF3E2"},
    ])
    return wrap_slide(s)


def slide_proposal_standard():
    s = ""
    s += _rect(2, "h", 0, 0, SLIDE_WIDTH, inches(1.1), NAVY)
    s += _textbox(3, "ht", inches(0.5), inches(0.15), inches(10), inches(0.9), [
        {"text": "10 | STANDARD PLAN - INR 10,000/month", "size": 1900, "bold": True, "color": WHITE},
    ])
    s += _rounded_rect(4, "plan", inches(0.4), inches(1.4), inches(7.2), inches(5.7), LIGHT_BG, [
        {"text": "STANDARD PLAN", "size": 1800, "bold": True, "color": NAVY},
        {"text": "Social Media Management Package", "size": 1100, "color": GREY},
        {"text": "", "size": 250, "color": TEXT},
        {"text": "INCLUDED SERVICES:", "size": 1100, "bold": True, "color": RED},
        {"text": "", "size": 150, "color": TEXT},
        {"text": "Monthly Content Calendar [8-11 Posts]", "size": 1050, "color": TEXT},
        {"text": "   1-2 Reels", "size": 950, "color": SUBTEXT},
        {"text": "   2-3 Carousels", "size": 950, "color": SUBTEXT},
        {"text": "   4-6 Static Posts", "size": 950, "color": SUBTEXT},
        {"text": "", "size": 100, "color": TEXT},
        {"text": "Hashtag Research", "size": 1050, "color": TEXT},
        {"text": "Basic SEO-Based Optimization", "size": 1050, "color": TEXT},
        {"text": "Basic Page Optimization", "size": 1050, "color": TEXT},
        {"text": "Monthly Performance Report", "size": 1050, "color": TEXT},
        {"text": "", "size": 250, "color": TEXT},
        {"text": "Platform: Facebook & Instagram Only", "size": 1050, "bold": True, "color": DARK},
        {"text": "NO Ads Management Included", "size": 1050, "color": RED},
    ])
    s += _rounded_rect(5, "price", inches(7.9), inches(1.4), inches(3.7), inches(3.0), NAVY, [
        {"text": "INVESTMENT", "size": 1200, "bold": True, "color": RED},
        {"text": "", "size": 300, "color": WHITE},
        {"text": "INR 10,000", "size": 2000, "bold": True, "color": WHITE},
        {"text": "/month", "size": 1200, "color": "BDC3C7"},
        {"text": "", "size": 250, "color": WHITE},
        {"text": "Best for businesses starting", "size": 1000, "color": "BDC3C7"},
        {"text": "their social media journey", "size": 1000, "color": "BDC3C7"},
    ])
    s += _rounded_rect(6, "ideal", inches(7.9), inches(4.6), inches(3.7), inches(2.5), LIGHT_BG, [
        {"text": "IDEAL FOR:", "size": 1050, "bold": True, "color": NAVY},
        {"text": "", "size": 150, "color": TEXT},
        {"text": "Businesses needing basic", "size": 950, "color": SUBTEXT},
        {"text": "social presence", "size": 950, "color": SUBTEXT},
        {"text": "Consistent posting without", "size": 950, "color": SUBTEXT},
        {"text": "ad spend commitment", "size": 950, "color": SUBTEXT},
        {"text": "Brand building on a budget", "size": 950, "color": SUBTEXT},
    ])
    return wrap_slide(s)


def slide_proposal_business():
    s = ""
    s += _rect(2, "h", 0, 0, SLIDE_WIDTH, inches(1.1), NAVY)
    s += _textbox(3, "ht", inches(0.5), inches(0.15), inches(10), inches(0.9), [
        {"text": "11 | BUSINESS PLAN - INR 18,000/month", "size": 1900, "bold": True, "color": WHITE},
    ])
    s += _rounded_rect(4, "plan", inches(0.4), inches(1.4), inches(7.2), inches(5.7), LIGHT_BG, [
        {"text": "BUSINESS PLAN", "size": 1800, "bold": True, "color": NAVY},
        {"text": "Complete Digital Growth Package", "size": 1100, "color": GREY},
        {"text": "", "size": 250, "color": TEXT},
        {"text": "INCLUDED SERVICES:", "size": 1100, "bold": True, "color": RED},
        {"text": "", "size": 150, "color": TEXT},
        {"text": "Monthly Content Calendar [18-20 Posts]", "size": 1050, "color": TEXT},
        {"text": "   2-4 Reels", "size": 950, "color": SUBTEXT},
        {"text": "   3-5 Carousels", "size": 950, "color": SUBTEXT},
        {"text": "   9-12 Static Posts", "size": 950, "color": SUBTEXT},
        {"text": "", "size": 100, "color": TEXT},
        {"text": "Platforms: Facebook, Instagram, LinkedIn, YouTube", "size": 1050, "bold": True, "color": DARK},
        {"text": "", "size": 100, "color": TEXT},
        {"text": "ADs Management Included", "size": 1050, "color": TEXT},
        {"text": "   (Budget up to $300 USD/month)", "size": 950, "color": SUBTEXT},
        {"text": "Basic GMB Optimization & Updates", "size": 1050, "color": TEXT},
        {"text": "Bi-Weekly Strategy Calls", "size": 1050, "color": TEXT},
        {"text": "Monthly Performance Report", "size": 1050, "color": TEXT},
    ])
    s += _rounded_rect(5, "price", inches(7.9), inches(1.4), inches(3.7), inches(3.2), NAVY, [
        {"text": "INVESTMENT", "size": 1200, "bold": True, "color": RED},
        {"text": "", "size": 300, "color": WHITE},
        {"text": "INR 18,000", "size": 2000, "bold": True, "color": WHITE},
        {"text": "/month", "size": 1200, "color": "BDC3C7"},
        {"text": "", "size": 250, "color": WHITE},
        {"text": "Ad budget separate:", "size": 1000, "color": "BDC3C7"},
        {"text": "$300 USD/month", "size": 1100, "bold": True, "color": WHITE},
        {"text": "(client-funded)", "size": 900, "color": "BDC3C7"},
    ])
    s += _rounded_rect(6, "reco", inches(7.9), inches(4.8), inches(3.7), inches(2.3), RED, [
        {"text": "RECOMMENDED", "size": 1100, "bold": True, "color": WHITE},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Best value for businesses", "size": 950, "color": WHITE},
        {"text": "ready to grow aggressively", "size": 950, "color": WHITE},
        {"text": "", "size": 100, "color": WHITE},
        {"text": "Includes ads management", "size": 950, "color": WHITE},
        {"text": "Multi-platform presence", "size": 950, "color": WHITE},
        {"text": "GMB + Strategy calls", "size": 950, "color": WHITE},
    ])
    return wrap_slide(s)


def slide_thank_you():
    s = ""
    s += _rect(2, "top", 0, 0, SLIDE_WIDTH, inches(0.12), RED)
    s += _textbox(3, "ty", inches(2), inches(1.8), inches(8), inches(2.2), [
        {"text": "READY TO GROW?", "size": 3600, "bold": True, "color": WHITE, "align": "ctr"},
        {"text": "Let's Build Something Great Together.", "size": 1600, "color": "BDC3C7", "align": "ctr", "spc_before": 400},
    ])
    s += _rounded_rect(4, "ns", inches(3), inches(4.0), inches(6), inches(2.8), DARK, [
        {"text": "NEXT STEPS", "size": 1300, "bold": True, "color": RED, "align": "ctr"},
        {"text": "", "size": 200, "color": WHITE},
        {"text": "1. Review this proposal", "size": 1100, "color": WHITE, "align": "ctr"},
        {"text": "2. Choose your plan (Standard / Business)", "size": 1100, "color": WHITE, "align": "ctr"},
        {"text": "3. Schedule a strategy call", "size": 1100, "color": WHITE, "align": "ctr"},
        {"text": "4. We start working within 48 hours", "size": 1100, "color": WHITE, "align": "ctr"},
    ])
    s += _rect(5, "bot", 0, inches(7.3), SLIDE_WIDTH, inches(0.2), RED)
    return wrap_slide(s, NAVY)
