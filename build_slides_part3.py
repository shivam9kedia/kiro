"""Slides 11-15: Ads Strategy, Roadmap, KPIs, Proposal, Thank You."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx_builder.core import *
from pptx_builder.slides import _textbox, _rect, _rounded_rect, wrap_slide


def slide_ads_strategy():
    """Slide 11: Ads Strategy."""
    shapes = ""
    shapes += _rect(2, "header_bg", 0, 0, SLIDE_WIDTH, inches(1.2), "1B2A4A")
    shapes += _textbox(3, "header", inches(0.5), inches(0.2), inches(10), inches(0.9), [
        {"text": "07 | ADS STRATEGY - ENGAGEMENT & FOLLOWER GROWTH", "size": 2000, "bold": True, "color": "FFFFFF", "align": "l"},
    ])
    # Facebook/Instagram Ads
    shapes += _rounded_rect(4, "engagement", inches(0.5), inches(1.5), inches(5.5), inches(3.0), "3498DB", [
        {"text": "ENGAGEMENT ADS", "size": 1200, "bold": True, "color": "FFFFFF"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Objective: Post Engagement", "size": 1000, "color": "FFFFFF"},
        {"text": "Budget: $300-500/month", "size": 1000, "color": "FFFFFF"},
        {"text": "Audience: Calgary homeowners 30-65", "size": 1000, "color": "FFFFFF"},
        {"text": "Interests: Home improvement, renovation", "size": 1000, "color": "FFFFFF"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Best Performing Ad Types:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "Before/After carousels", "size": 950, "color": "D6EAF8"},
        {"text": "Video testimonials (15-30 sec)", "size": 950, "color": "D6EAF8"},
        {"text": "Seasonal tips with CTA", "size": 950, "color": "D6EAF8"},
    ])
    # Follower Growth Ads
    shapes += _rounded_rect(5, "followers", inches(0.5), inches(4.7), inches(5.5), inches(2.5), "2ECC71", [
        {"text": "FOLLOWER GROWTH ADS", "size": 1200, "bold": True, "color": "FFFFFF"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Objective: Page Likes / Profile Visits", "size": 1000, "color": "FFFFFF"},
        {"text": "Budget: $200-300/month", "size": 1000, "color": "FFFFFF"},
        {"text": "Targeting: Lookalike of current customers", "size": 1000, "color": "FFFFFF"},
        {"text": "Calgary + surrounding 50km radius", "size": 1000, "color": "FFFFFF"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Realistic Goal: +150-300 followers/month", "size": 1000, "bold": True, "color": "FFFFFF"},
    ])
    # Ad calendar
    shapes += _rounded_rect(6, "ad_cal", inches(6.3), inches(1.5), inches(5.3), inches(5.7), "F0F4F8", [
        {"text": "3-MONTH AD CALENDAR", "size": 1200, "bold": True, "color": "1B2A4A"},
        {"text": "", "size": 300, "color": "333333"},
        {"text": "MONTH 1:", "size": 1050, "bold": True, "color": "E74C3C"},
        {"text": "Brand awareness + page likes", "size": 950, "color": "555555"},
        {"text": "Boost top 2 posts weekly", "size": 950, "color": "555555"},
        {"text": "Budget: $400", "size": 950, "color": "555555"},
        {"text": "", "size": 200, "color": "333333"},
        {"text": "MONTH 2:", "size": 1050, "bold": True, "color": "F39C12"},
        {"text": "Engagement + video views", "size": 950, "color": "555555"},
        {"text": "Boost Reels + testimonials", "size": 950, "color": "555555"},
        {"text": "Budget: $500", "size": 950, "color": "555555"},
        {"text": "", "size": 200, "color": "333333"},
        {"text": "MONTH 3:", "size": 1050, "bold": True, "color": "2ECC71"},
        {"text": "Leads + retargeting warm audience", "size": 950, "color": "555555"},
        {"text": "Carousel ads + offer promos", "size": 950, "color": "555555"},
        {"text": "Budget: $600", "size": 950, "color": "555555"},
        {"text": "", "size": 300, "color": "333333"},
        {"text": "TOTAL 3-MONTH AD SPEND: $1,500", "size": 1100, "bold": True, "color": "1B2A4A"},
    ])
    return wrap_slide(shapes)


def slide_roadmap_overview():
    """Slide 12: Month-by-Month Roadmap."""
    shapes = ""
    shapes += _rect(2, "header_bg", 0, 0, SLIDE_WIDTH, inches(1.2), "1B2A4A")
    shapes += _textbox(3, "header", inches(0.5), inches(0.2), inches(10), inches(0.9), [
        {"text": "08 | MONTH-BY-MONTH ROADMAP", "size": 2000, "bold": True, "color": "FFFFFF", "align": "l"},
    ])
    # Month 1
    shapes += _rounded_rect(4, "m1", inches(0.3), inches(1.5), inches(3.7), inches(5.7), "E74C3C", [
        {"text": "MONTH 1", "size": 1400, "bold": True, "color": "FFFFFF"},
        {"text": "FOUNDATION", "size": 1100, "color": "FFDDD5"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Website Fixes:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "Meta tags optimization", "size": 900, "color": "FFDDD5"},
        {"text": "Schema markup added", "size": 900, "color": "FFDDD5"},
        {"text": "Speed optimization", "size": 900, "color": "FFDDD5"},
        {"text": "", "size": 150, "color": "FFFFFF"},
        {"text": "Social Media:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "Profile optimization", "size": 900, "color": "FFDDD5"},
        {"text": "Content pillars set", "size": 900, "color": "FFDDD5"},
        {"text": "3-4 posts/week start", "size": 900, "color": "FFDDD5"},
        {"text": "", "size": 150, "color": "FFFFFF"},
        {"text": "GMB:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "Full optimization", "size": 900, "color": "FFDDD5"},
        {"text": "Weekly posts begin", "size": 900, "color": "FFDDD5"},
        {"text": "Review campaign start", "size": 900, "color": "FFDDD5"},
    ])
    # Month 2
    shapes += _rounded_rect(5, "m2", inches(4.15), inches(1.5), inches(3.7), inches(5.7), "F39C12", [
        {"text": "MONTH 2", "size": 1400, "bold": True, "color": "FFFFFF"},
        {"text": "ENGAGEMENT", "size": 1100, "color": "FEF3E2"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Website:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "Blog launches (2 posts)", "size": 900, "color": "FEF3E2"},
        {"text": "Location pages created", "size": 900, "color": "FEF3E2"},
        {"text": "Testimonials page", "size": 900, "color": "FEF3E2"},
        {"text": "", "size": 150, "color": "FFFFFF"},
        {"text": "Social Media:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "4-5 posts/week", "size": 900, "color": "FEF3E2"},
        {"text": "Reels strategy active", "size": 900, "color": "FEF3E2"},
        {"text": "Engagement campaigns", "size": 900, "color": "FEF3E2"},
        {"text": "", "size": 150, "color": "FFFFFF"},
        {"text": "Ads:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "Engagement ads live", "size": 900, "color": "FEF3E2"},
        {"text": "Video view campaigns", "size": 900, "color": "FEF3E2"},
        {"text": "A/B testing creatives", "size": 900, "color": "FEF3E2"},
    ])
    # Month 3
    shapes += _rounded_rect(6, "m3", inches(8.0), inches(1.5), inches(3.7), inches(5.7), "2ECC71", [
        {"text": "MONTH 3", "size": 1400, "bold": True, "color": "FFFFFF"},
        {"text": "CONVERSION", "size": 1100, "color": "D5F5E3"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Website:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "2 more blog posts", "size": 900, "color": "D5F5E3"},
        {"text": "CTA optimization", "size": 900, "color": "D5F5E3"},
        {"text": "Lead forms improved", "size": 900, "color": "D5F5E3"},
        {"text": "", "size": 150, "color": "FFFFFF"},
        {"text": "Social Media:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "Lead gen content", "size": 900, "color": "D5F5E3"},
        {"text": "Customer stories", "size": 900, "color": "D5F5E3"},
        {"text": "Conversion CTAs", "size": 900, "color": "D5F5E3"},
        {"text": "", "size": 150, "color": "FFFFFF"},
        {"text": "Ads:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "Retargeting active", "size": 900, "color": "D5F5E3"},
        {"text": "Lead gen campaigns", "size": 900, "color": "D5F5E3"},
        {"text": "Offer promotions", "size": 900, "color": "D5F5E3"},
    ])
    return wrap_slide(shapes)


def slide_kpis():
    """Slide 13: KPIs & Success Metrics."""
    shapes = ""
    shapes += _rect(2, "header_bg", 0, 0, SLIDE_WIDTH, inches(1.2), "1B2A4A")
    shapes += _textbox(3, "header", inches(0.5), inches(0.2), inches(10), inches(0.9), [
        {"text": "09 | KPIs & REALISTIC SUCCESS METRICS", "size": 2000, "bold": True, "color": "FFFFFF", "align": "l"},
    ])
    shapes += _textbox(4, "note", inches(0.5), inches(1.3), inches(11), inches(0.4), [
        {"text": "No fake promises. These are achievable targets based on industry benchmarks for local service businesses.", "size": 1050, "color": "7F8C8D"},
    ])
    # Social metrics
    shapes += _rounded_rect(5, "social_kpi", inches(0.5), inches(1.9), inches(3.5), inches(5.2), "3498DB", [
        {"text": "SOCIAL MEDIA", "size": 1200, "bold": True, "color": "FFFFFF"},
        {"text": "", "size": 300, "color": "FFFFFF"},
        {"text": "Instagram Followers:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "+300-500 in 3 months", "size": 1000, "color": "D6EAF8"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Engagement Rate:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "Target 3-5% (industry avg 1-2%)", "size": 1000, "color": "D6EAF8"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Facebook Page Likes:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "+200-400 in 3 months", "size": 1000, "color": "D6EAF8"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Post Reach:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "2x-3x current reach", "size": 1000, "color": "D6EAF8"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Reels Views:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "500-2,000 per Reel avg", "size": 1000, "color": "D6EAF8"},
    ])
    # Website metrics
    shapes += _rounded_rect(6, "web_kpi", inches(4.2), inches(1.9), inches(3.5), inches(5.2), "2ECC71", [
        {"text": "WEBSITE & SEO", "size": 1200, "bold": True, "color": "FFFFFF"},
        {"text": "", "size": 300, "color": "FFFFFF"},
        {"text": "Organic Traffic:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "+30-50% increase", "size": 1000, "color": "D5F5E3"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Keyword Rankings:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "Top 20 for 5+ local terms", "size": 1000, "color": "D5F5E3"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "GMB Views:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "+40-60% increase", "size": 1000, "color": "D5F5E3"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Google Reviews:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "+10-15 new reviews", "size": 1000, "color": "D5F5E3"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Website Leads:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "+20-40% more inquiries", "size": 1000, "color": "D5F5E3"},
    ])
    # Business metrics
    shapes += _rounded_rect(7, "biz_kpi", inches(7.9), inches(1.9), inches(3.7), inches(5.2), "F39C12", [
        {"text": "BUSINESS IMPACT", "size": 1200, "bold": True, "color": "FFFFFF"},
        {"text": "", "size": 300, "color": "FFFFFF"},
        {"text": "Quote Requests:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "+15-25% from digital", "size": 1000, "color": "FEF3E2"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Brand Awareness:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "Measurable via reach & mentions", "size": 1000, "color": "FEF3E2"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Cost Per Lead:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "Target: $15-30 per lead", "size": 1000, "color": "FEF3E2"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Review Rating:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "Maintain 4.5+ stars", "size": 1000, "color": "FEF3E2"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Community Growth:", "size": 1000, "bold": True, "color": "FFFFFF"},
        {"text": "Engaged local following", "size": 1000, "color": "FEF3E2"},
    ])
    return wrap_slide(shapes)


def slide_proposal():
    """Slide 14: Standard Plan Proposal."""
    shapes = ""
    shapes += _rect(2, "header_bg", 0, 0, SLIDE_WIDTH, inches(1.2), "1B2A4A")
    shapes += _textbox(3, "header", inches(0.5), inches(0.2), inches(10), inches(0.9), [
        {"text": "10 | STANDARD PLAN - PROPOSAL", "size": 2000, "bold": True, "color": "FFFFFF", "align": "l"},
    ])
    # Standard Plan
    shapes += _rounded_rect(4, "plan", inches(0.5), inches(1.5), inches(7.0), inches(5.7), "F0F4F8", [
        {"text": "STANDARD PLAN", "size": 1800, "bold": True, "color": "1B2A4A"},
        {"text": "3-Month Digital Growth Package", "size": 1100, "color": "7F8C8D"},
        {"text": "", "size": 300, "color": "333333"},
        {"text": "INCLUDED SERVICES:", "size": 1100, "bold": True, "color": "E74C3C"},
        {"text": "", "size": 150, "color": "333333"},
        {"text": "Social Media Management (IG + FB)", "size": 1000, "color": "333333"},
        {"text": "   12-16 posts/month + Stories", "size": 900, "color": "555555"},
        {"text": "Content Creation (Graphics + Reels)", "size": 1000, "color": "333333"},
        {"text": "   Professional visuals + short-form video", "size": 900, "color": "555555"},
        {"text": "Community Management & Engagement", "size": 1000, "color": "333333"},
        {"text": "   Daily monitoring + responses", "size": 900, "color": "555555"},
        {"text": "Website SEO Optimization (On-Page)", "size": 1000, "color": "333333"},
        {"text": "   Meta tags, schema, speed fixes", "size": 900, "color": "555555"},
        {"text": "Blog Content (2 SEO articles/month)", "size": 1000, "color": "333333"},
        {"text": "GMB Management & Weekly Posts", "size": 1000, "color": "333333"},
        {"text": "Monthly Analytics Report & Strategy Call", "size": 1000, "color": "333333"},
        {"text": "Ad Management (budget separate)", "size": 1000, "color": "333333"},
    ])
    # Pricing box
    shapes += _rounded_rect(5, "pricing", inches(7.8), inches(1.5), inches(3.8), inches(3.0), "1B2A4A", [
        {"text": "INVESTMENT", "size": 1200, "bold": True, "color": "E74C3C"},
        {"text": "", "size": 300, "color": "FFFFFF"},
        {"text": "$______/month", "size": 1600, "bold": True, "color": "FFFFFF"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "3-month commitment", "size": 1000, "color": "BDC3C7"},
        {"text": "Ad budget separate", "size": 1000, "color": "BDC3C7"},
        {"text": "(recommended $500-600/mo)", "size": 900, "color": "BDC3C7"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Setup fee: $______", "size": 1000, "color": "BDC3C7"},
        {"text": "(one-time, Month 1 only)", "size": 900, "color": "BDC3C7"},
    ])
    # What's NOT included
    shapes += _rounded_rect(6, "not_inc", inches(7.8), inches(4.7), inches(3.8), inches(2.5), "F8F0F0", [
        {"text": "NOTE:", "size": 1000, "bold": True, "color": "E74C3C"},
        {"text": "", "size": 150, "color": "333333"},
        {"text": "Ad spend is client-funded", "size": 900, "color": "555555"},
        {"text": "Website redesign not included", "size": 900, "color": "555555"},
        {"text": "Photography/video shoots extra", "size": 900, "color": "555555"},
        {"text": "Pricing shown is placeholder -", "size": 900, "color": "555555"},
        {"text": "customize before presenting", "size": 900, "color": "555555"},
    ])
    return wrap_slide(shapes)


def slide_thank_you():
    """Slide 15: Thank You / Next Steps."""
    shapes = ""
    # Accent bar
    shapes += _rect(2, "top_bar", 0, 0, SLIDE_WIDTH, inches(0.15), "E74C3C")
    # Main text
    shapes += _textbox(3, "thanks", inches(2), inches(2.0), inches(8), inches(2), [
        {"text": "READY TO GROW?", "size": 3600, "bold": True, "color": "FFFFFF", "align": "ctr"},
        {"text": "Let's Build Something Great Together.", "size": 1600, "color": "BDC3C7", "align": "ctr", "spc_before": 400},
    ])
    # Next steps
    shapes += _rounded_rect(4, "next", inches(3), inches(4.2), inches(6), inches(2.5), "2C3E50", [
        {"text": "NEXT STEPS", "size": 1300, "bold": True, "color": "E74C3C", "align": "ctr"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "1. Review this proposal", "size": 1100, "color": "FFFFFF", "align": "ctr"},
        {"text": "2. Schedule a strategy call", "size": 1100, "color": "FFFFFF", "align": "ctr"},
        {"text": "3. Confirm plan & timeline", "size": 1100, "color": "FFFFFF", "align": "ctr"},
        {"text": "4. We start working within 48 hours", "size": 1100, "color": "FFFFFF", "align": "ctr"},
    ])
    # Bottom bar
    shapes += _rect(5, "bottom_bar", 0, inches(7.2), SLIDE_WIDTH, inches(0.3), "E74C3C")
    return wrap_slide(shapes, "1B2A4A")
