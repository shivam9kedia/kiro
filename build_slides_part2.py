"""Slides 5-9: Competitor Analysis, Social Audit, Opportunities, Content Plans."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx_builder.core import *
from pptx_builder.slides import _textbox, _rect, _rounded_rect, wrap_slide


def slide_competitor_analysis():
    """Slide 5: Competitor Analysis."""
    shapes = ""
    shapes += _rect(2, "header_bg", 0, 0, SLIDE_WIDTH, inches(1.2), "1B2A4A")
    shapes += _textbox(3, "header", inches(0.5), inches(0.2), inches(10), inches(0.9), [
        {"text": "03 | COMPETITOR ANALYSIS", "size": 2000, "bold": True, "color": "FFFFFF", "align": "l"},
    ])
    # Competitor comparison
    shapes += _textbox(4, "subtitle", inches(0.5), inches(1.4), inches(11), inches(0.5), [
        {"text": "Key Competitors & What They Do Better", "size": 1300, "bold": True, "color": "2C3E50"},
    ])
    # Competitor cards
    competitors = [
        ("Lux Windows", "3498DB", "Strong brand identity, premium positioning, active social media, video content, showroom experience"),
        ("All Weather", "2ECC71", "10x Energy Star winner, extensive product range, strong blog/SEO, dealer network, Calgary + Edmonton"),
        ("Durabuilt", "F39C12", "Multi-city presence, design galleries, active social channels, builder partnerships, warranty focus"),
    ]
    y_pos = inches(2.1)
    for name, color, desc in competitors:
        shapes += _rounded_rect(10, f"comp_{name}", inches(0.5), y_pos, inches(3.5), inches(1.5), color, [
            {"text": name, "size": 1200, "bold": True, "color": "FFFFFF"},
            {"text": desc, "size": 850, "color": "FFFFFF"},
        ])
        y_pos += inches(1.7)

    competitors2 = [
        ("JELD-WEN", "9B59B6", "National brand power, massive product catalog, strong SEO domain authority, content marketing machine"),
        ("WindowMart", "1ABC9C", "Aggressive pricing, strong Google Ads, multiple locations, quick quotes, review generation"),
        ("Windows Canada", "E67E22", "Modern website, project galleries, financing options, active blog, good local SEO structure"),
    ]
    y_pos = inches(2.1)
    for name, color, desc in competitors2:
        shapes += _rounded_rect(11, f"comp2_{name}", inches(4.2), y_pos, inches(3.5), inches(1.5), color, [
            {"text": name, "size": 1200, "bold": True, "color": "FFFFFF"},
            {"text": desc, "size": 850, "color": "FFFFFF"},
        ])
        y_pos += inches(1.7)

    # Key takeaway
    shapes += _rounded_rect(20, "takeaway", inches(8.0), inches(2.1), inches(3.7), inches(4.8), "1B2A4A", [
        {"text": "KEY GAPS", "size": 1100, "bold": True, "color": "E74C3C"},
        {"text": "", "size": 300, "color": "FFFFFF"},
        {"text": "Apollo lacks:", "size": 1000, "color": "FFFFFF"},
        {"text": "Active blog content", "size": 900, "color": "BDC3C7"},
        {"text": "Video/Reels presence", "size": 900, "color": "BDC3C7"},
        {"text": "Project galleries", "size": 900, "color": "BDC3C7"},
        {"text": "Review generation", "size": 900, "color": "BDC3C7"},
        {"text": "Social engagement", "size": 900, "color": "BDC3C7"},
        {"text": "Paid ads strategy", "size": 900, "color": "BDC3C7"},
        {"text": "", "size": 300, "color": "FFFFFF"},
        {"text": "OPPORTUNITY:", "size": 1000, "bold": True, "color": "2ECC71"},
        {"text": "Local trust + family story is a differentiator", "size": 900, "color": "FFFFFF"},
    ])
    return wrap_slide(shapes)


def slide_social_audit():
    """Slide 6: Social Media Audit."""
    shapes = ""
    shapes += _rect(2, "header_bg", 0, 0, SLIDE_WIDTH, inches(1.2), "1B2A4A")
    shapes += _textbox(3, "header", inches(0.5), inches(0.2), inches(10), inches(0.9), [
        {"text": "04 | SOCIAL MEDIA AUDIT", "size": 2000, "bold": True, "color": "FFFFFF", "align": "l"},
    ])
    # Instagram audit
    shapes += _rounded_rect(4, "ig_box", inches(0.5), inches(1.5), inches(5.5), inches(5.5), "F8F0FC", [
        {"text": "INSTAGRAM @apollowindowscalgary", "size": 1200, "bold": True, "color": "9B59B6"},
        {"text": "", "size": 300, "color": "333333"},
        {"text": "Current Status:", "size": 1100, "bold": True, "color": "333333"},
        {"text": "Low follower count & engagement", "size": 1000, "color": "555555"},
        {"text": "Inconsistent posting frequency", "size": 1000, "color": "555555"},
        {"text": "Limited use of Reels/Stories", "size": 1000, "color": "555555"},
        {"text": "No hashtag strategy", "size": 1000, "color": "555555"},
        {"text": "Missing bio link optimization", "size": 1000, "color": "555555"},
        {"text": "", "size": 300, "color": "333333"},
        {"text": "Issues:", "size": 1100, "bold": True, "color": "E74C3C"},
        {"text": "No content pillars defined", "size": 1000, "color": "555555"},
        {"text": "No community engagement", "size": 1000, "color": "555555"},
        {"text": "Product-only posts (no storytelling)", "size": 1000, "color": "555555"},
        {"text": "No user-generated content", "size": 1000, "color": "555555"},
    ])
    # Facebook audit
    shapes += _rounded_rect(5, "fb_box", inches(6.3), inches(1.5), inches(5.3), inches(5.5), "F0F4FF", [
        {"text": "FACEBOOK @apollowindowscalgary", "size": 1200, "bold": True, "color": "3498DB"},
        {"text": "", "size": 300, "color": "333333"},
        {"text": "Current Status:", "size": 1100, "bold": True, "color": "333333"},
        {"text": "Page exists but underutilized", "size": 1000, "color": "555555"},
        {"text": "Low engagement rate on posts", "size": 1000, "color": "555555"},
        {"text": "Few customer reviews on page", "size": 1000, "color": "555555"},
        {"text": "No Facebook Events or Offers used", "size": 1000, "color": "555555"},
        {"text": "Limited community interaction", "size": 1000, "color": "555555"},
        {"text": "", "size": 300, "color": "333333"},
        {"text": "Issues:", "size": 1100, "bold": True, "color": "E74C3C"},
        {"text": "No paid promotion strategy", "size": 1000, "color": "555555"},
        {"text": "Missing Facebook Shop/catalog", "size": 1000, "color": "555555"},
        {"text": "No video content or Lives", "size": 1000, "color": "555555"},
        {"text": "Inconsistent brand voice", "size": 1000, "color": "555555"},
    ])
    return wrap_slide(shapes)


def slide_opportunities():
    """Slide 7: Opportunities & Quick Wins."""
    shapes = ""
    shapes += _rect(2, "header_bg", 0, 0, SLIDE_WIDTH, inches(1.2), "1B2A4A")
    shapes += _textbox(3, "header", inches(0.5), inches(0.2), inches(10), inches(0.9), [
        {"text": "05 | OPPORTUNITIES & QUICK WINS", "size": 2000, "bold": True, "color": "FFFFFF", "align": "l"},
    ])
    # Quick wins
    shapes += _rounded_rect(4, "qw", inches(0.5), inches(1.5), inches(5.5), inches(3.0), "2ECC71", [
        {"text": "QUICK WINS (Week 1-2)", "size": 1300, "bold": True, "color": "FFFFFF"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Optimize GMB profile completely", "size": 1000, "color": "FFFFFF"},
        {"text": "Fix website meta titles & descriptions", "size": 1000, "color": "FFFFFF"},
        {"text": "Set up consistent posting schedule", "size": 1000, "color": "FFFFFF"},
        {"text": "Create branded hashtag set", "size": 1000, "color": "FFFFFF"},
        {"text": "Ask 10 past clients for Google reviews", "size": 1000, "color": "FFFFFF"},
        {"text": "Update Instagram bio + link tree", "size": 1000, "color": "FFFFFF"},
    ])
    # Social media opportunities
    shapes += _rounded_rect(5, "sm_opp", inches(0.5), inches(4.7), inches(5.5), inches(2.5), "3498DB", [
        {"text": "SOCIAL MEDIA OPPORTUNITIES", "size": 1200, "bold": True, "color": "FFFFFF"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Before/After project reveals (high engagement)", "size": 1000, "color": "FFFFFF"},
        {"text": "Behind-the-scenes manufacturing Reels", "size": 1000, "color": "FFFFFF"},
        {"text": "Calgary homeowner tips (seasonal)", "size": 1000, "color": "FFFFFF"},
        {"text": "Customer testimonial videos", "size": 1000, "color": "FFFFFF"},
        {"text": "Local partnership cross-promos", "size": 1000, "color": "FFFFFF"},
    ])
    # Growth opportunities
    shapes += _rounded_rect(6, "growth", inches(6.3), inches(1.5), inches(5.3), inches(5.7), "F39C12", [
        {"text": "GROWTH OPPORTUNITIES", "size": 1300, "bold": True, "color": "FFFFFF"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "LOCAL SEO:", "size": 1100, "bold": True, "color": "FFFFFF"},
        {"text": "Rank for 'windows Calgary' keywords", "size": 1000, "color": "FFFFFF"},
        {"text": "Target neighbourhood-specific searches", "size": 1000, "color": "FFFFFF"},
        {"text": "Build local citation consistency", "size": 1000, "color": "FFFFFF"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "CONTENT:", "size": 1100, "bold": True, "color": "FFFFFF"},
        {"text": "Energy efficiency educational content", "size": 1000, "color": "FFFFFF"},
        {"text": "Seasonal maintenance guides", "size": 1000, "color": "FFFFFF"},
        {"text": "Product comparison guides", "size": 1000, "color": "FFFFFF"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "PARTNERSHIPS:", "size": 1100, "bold": True, "color": "FFFFFF"},
        {"text": "Local builders & contractors", "size": 1000, "color": "FFFFFF"},
        {"text": "Home renovation influencers", "size": 1000, "color": "FFFFFF"},
        {"text": "Calgary real estate agents", "size": 1000, "color": "FFFFFF"},
    ])
    return wrap_slide(shapes)


def slide_content_plan_month1():
    """Slide 8: Content Plan - Month 1."""
    shapes = ""
    shapes += _rect(2, "header_bg", 0, 0, SLIDE_WIDTH, inches(1.2), "1B2A4A")
    shapes += _textbox(3, "header", inches(0.5), inches(0.2), inches(10), inches(0.9), [
        {"text": "06 | CONTENT PLAN - MONTH 1 (Foundation)", "size": 2000, "bold": True, "color": "FFFFFF", "align": "l"},
    ])
    shapes += _textbox(4, "focus", inches(0.5), inches(1.3), inches(11), inches(0.5), [
        {"text": "THEME: Build Foundation & Brand Awareness | 3-4 posts/week", "size": 1200, "bold": True, "color": "E74C3C"},
    ])
    # Week 1-2
    shapes += _rounded_rect(5, "w1", inches(0.5), inches(1.9), inches(5.5), inches(2.5), "F0F4F8", [
        {"text": "WEEK 1-2: BRAND STORY", "size": 1100, "bold": True, "color": "1B2A4A"},
        {"text": "", "size": 200, "color": "333333"},
        {"text": "Mon: Meet the Team (carousel)", "size": 950, "color": "333333"},
        {"text": "Wed: Our Manufacturing Process (Reel)", "size": 950, "color": "333333"},
        {"text": "Fri: Before/After Project #1 (carousel)", "size": 950, "color": "333333"},
        {"text": "Sat: Customer Testimonial (Story)", "size": 950, "color": "333333"},
        {"text": "", "size": 200, "color": "333333"},
        {"text": "Blog: 'Why Choose Local Windows Manufacturer'", "size": 950, "color": "3498DB"},
    ])
    # Week 3-4
    shapes += _rounded_rect(6, "w2", inches(0.5), inches(4.6), inches(5.5), inches(2.5), "F0F4F8", [
        {"text": "WEEK 3-4: EDUCATION & TRUST", "size": 1100, "bold": True, "color": "1B2A4A"},
        {"text": "", "size": 200, "color": "333333"},
        {"text": "Mon: Energy Efficiency Tips (infographic)", "size": 950, "color": "333333"},
        {"text": "Wed: Window Style Guide (carousel)", "size": 950, "color": "333333"},
        {"text": "Fri: Before/After Project #2 (Reel)", "size": 950, "color": "333333"},
        {"text": "Sat: FAQ Answer (Story poll)", "size": 950, "color": "333333"},
        {"text": "", "size": 200, "color": "333333"},
        {"text": "Blog: 'Calgary Winter Window Prep Guide'", "size": 950, "color": "3498DB"},
    ])
    # Content pillars
    shapes += _rounded_rect(7, "pillars", inches(6.3), inches(1.9), inches(5.3), inches(5.2), "1B2A4A", [
        {"text": "CONTENT PILLARS", "size": 1200, "bold": True, "color": "E74C3C"},
        {"text": "", "size": 300, "color": "FFFFFF"},
        {"text": "1. Project Showcases (30%)", "size": 1050, "color": "FFFFFF"},
        {"text": "   Before/after, time-lapses", "size": 900, "color": "BDC3C7"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "2. Educational (25%)", "size": 1050, "color": "FFFFFF"},
        {"text": "   Tips, guides, how-tos", "size": 900, "color": "BDC3C7"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "3. Behind the Scenes (20%)", "size": 1050, "color": "FFFFFF"},
        {"text": "   Manufacturing, team, process", "size": 900, "color": "BDC3C7"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "4. Social Proof (15%)", "size": 1050, "color": "FFFFFF"},
        {"text": "   Reviews, testimonials, awards", "size": 900, "color": "BDC3C7"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "5. Community (10%)", "size": 1050, "color": "FFFFFF"},
        {"text": "   Local events, partnerships", "size": 900, "color": "BDC3C7"},
    ])
    return wrap_slide(shapes)


def slide_content_plan_month2():
    """Slide 9: Content Plan - Month 2."""
    shapes = ""
    shapes += _rect(2, "header_bg", 0, 0, SLIDE_WIDTH, inches(1.2), "1B2A4A")
    shapes += _textbox(3, "header", inches(0.5), inches(0.2), inches(10), inches(0.9), [
        {"text": "06 | CONTENT PLAN - MONTH 2 (Engagement)", "size": 2000, "bold": True, "color": "FFFFFF", "align": "l"},
    ])
    shapes += _textbox(4, "focus", inches(0.5), inches(1.3), inches(11), inches(0.5), [
        {"text": "THEME: Drive Engagement & Community Building | 4-5 posts/week", "size": 1200, "bold": True, "color": "E74C3C"},
    ])
    # Week 5-6
    shapes += _rounded_rect(5, "w5", inches(0.5), inches(1.9), inches(5.5), inches(2.5), "F0F4F8", [
        {"text": "WEEK 5-6: ENGAGEMENT BOOST", "size": 1100, "bold": True, "color": "1B2A4A"},
        {"text": "", "size": 200, "color": "333333"},
        {"text": "Mon: 'Guess the Window Style' (poll)", "size": 950, "color": "333333"},
        {"text": "Tue: Installation Day Reel (timelapse)", "size": 950, "color": "333333"},
        {"text": "Thu: Homeowner Spotlight (interview)", "size": 950, "color": "333333"},
        {"text": "Fri: Product Close-up (carousel)", "size": 950, "color": "333333"},
        {"text": "Sat: Weekend DIY tip (Story)", "size": 950, "color": "333333"},
        {"text": "", "size": 200, "color": "333333"},
        {"text": "Blog: 'Triple vs Double Pane: Calgary Guide'", "size": 950, "color": "3498DB"},
    ])
    # Week 7-8
    shapes += _rounded_rect(6, "w7", inches(0.5), inches(4.6), inches(5.5), inches(2.5), "F0F4F8", [
        {"text": "WEEK 7-8: SEASONAL & LOCAL", "size": 1100, "bold": True, "color": "1B2A4A"},
        {"text": "", "size": 200, "color": "333333"},
        {"text": "Mon: 'Calgary Weather vs. Our Windows' (Reel)", "size": 950, "color": "333333"},
        {"text": "Wed: Local Builder Partnership Feature", "size": 950, "color": "333333"},
        {"text": "Thu: Energy Savings Calculator Post", "size": 950, "color": "333333"},
        {"text": "Fri: Project Reveal (Reel + carousel)", "size": 950, "color": "333333"},
        {"text": "Sat: Customer review highlight", "size": 950, "color": "333333"},
        {"text": "", "size": 200, "color": "333333"},
        {"text": "Blog: '5 Signs You Need Window Replacement'", "size": 950, "color": "3498DB"},
    ])
    # Creative ideas
    shapes += _rounded_rect(7, "creative", inches(6.3), inches(1.9), inches(5.3), inches(5.2), "9B59B6", [
        {"text": "CREATIVE CONTENT IDEAS", "size": 1200, "bold": True, "color": "FFFFFF"},
        {"text": "", "size": 300, "color": "FFFFFF"},
        {"text": "'Window Transformation Tuesday'", "size": 1050, "color": "FFFFFF"},
        {"text": "   Weekly series showing project reveals", "size": 900, "color": "E8D5F5"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "'Ask Apollo' Q&A Stories", "size": 1050, "color": "FFFFFF"},
        {"text": "   Answer common homeowner questions", "size": 900, "color": "E8D5F5"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "'Made in Calgary' Series", "size": 1050, "color": "FFFFFF"},
        {"text": "   Show local manufacturing pride", "size": 900, "color": "E8D5F5"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "'Energy Bill Challenge'", "size": 1050, "color": "FFFFFF"},
        {"text": "   Before/after energy savings", "size": 900, "color": "E8D5F5"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "'Meet Your Installer' Reels", "size": 1050, "color": "FFFFFF"},
        {"text": "   Humanize the brand", "size": 900, "color": "E8D5F5"},
    ])
    return wrap_slide(shapes)


def slide_content_plan_month3():
    """Slide 10: Content Plan - Month 3."""
    shapes = ""
    shapes += _rect(2, "header_bg", 0, 0, SLIDE_WIDTH, inches(1.2), "1B2A4A")
    shapes += _textbox(3, "header", inches(0.5), inches(0.2), inches(10), inches(0.9), [
        {"text": "06 | CONTENT PLAN - MONTH 3 (Conversion)", "size": 2000, "bold": True, "color": "FFFFFF", "align": "l"},
    ])
    shapes += _textbox(4, "focus", inches(0.5), inches(1.3), inches(11), inches(0.5), [
        {"text": "THEME: Convert Followers to Leads | 4-5 posts/week + Stories daily", "size": 1200, "bold": True, "color": "E74C3C"},
    ])
    # Week 9-10
    shapes += _rounded_rect(5, "w9", inches(0.5), inches(1.9), inches(5.5), inches(2.5), "F0F4F8", [
        {"text": "WEEK 9-10: SOCIAL PROOF & OFFERS", "size": 1100, "bold": True, "color": "1B2A4A"},
        {"text": "", "size": 200, "color": "333333"},
        {"text": "Mon: Customer Story Video (Reel)", "size": 950, "color": "333333"},
        {"text": "Tue: Limited-Time Offer Graphic", "size": 950, "color": "333333"},
        {"text": "Thu: '25 Years of Craftsmanship' carousel", "size": 950, "color": "333333"},
        {"text": "Fri: Side-by-side comparison (us vs generic)", "size": 950, "color": "333333"},
        {"text": "Sat: Free consultation CTA Story", "size": 950, "color": "333333"},
        {"text": "", "size": 200, "color": "333333"},
        {"text": "Blog: 'How to Choose the Right Windows for Your Home'", "size": 950, "color": "3498DB"},
    ])
    # Week 11-12
    shapes += _rounded_rect(6, "w11", inches(0.5), inches(4.6), inches(5.5), inches(2.5), "F0F4F8", [
        {"text": "WEEK 11-12: LEAD GENERATION", "size": 1100, "bold": True, "color": "1B2A4A"},
        {"text": "", "size": 200, "color": "333333"},
        {"text": "Mon: Project Portfolio Reel (best work)", "size": 950, "color": "333333"},
        {"text": "Wed: 'Why Apollo' comparison post", "size": 950, "color": "333333"},
        {"text": "Thu: Financing options infographic", "size": 950, "color": "333333"},
        {"text": "Fri: End-of-quarter results/celebration", "size": 950, "color": "333333"},
        {"text": "Sat: 'Book Your Free Quote' CTA", "size": 950, "color": "333333"},
        {"text": "", "size": 200, "color": "333333"},
        {"text": "Blog: 'Window ROI: What Calgary Homeowners Should Know'", "size": 950, "color": "3498DB"},
    ])
    # Conversion tactics
    shapes += _rounded_rect(7, "conversion", inches(6.3), inches(1.9), inches(5.3), inches(5.2), "E74C3C", [
        {"text": "CONVERSION TACTICS", "size": 1200, "bold": True, "color": "FFFFFF"},
        {"text": "", "size": 300, "color": "FFFFFF"},
        {"text": "Lead Magnets:", "size": 1050, "bold": True, "color": "FFFFFF"},
        {"text": "Free Window Style Guide PDF", "size": 950, "color": "FFDDD5"},
        {"text": "Energy Savings Calculator", "size": 950, "color": "FFDDD5"},
        {"text": "Free In-Home Consultation", "size": 950, "color": "FFDDD5"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "CTA Strategies:", "size": 1050, "bold": True, "color": "FFFFFF"},
        {"text": "Story swipe-up to quote form", "size": 950, "color": "FFDDD5"},
        {"text": "Bio link to booking page", "size": 950, "color": "FFDDD5"},
        {"text": "DM automation for inquiries", "size": 950, "color": "FFDDD5"},
        {"text": "", "size": 200, "color": "FFFFFF"},
        {"text": "Retargeting:", "size": 1050, "bold": True, "color": "FFFFFF"},
        {"text": "Website visitors > social ads", "size": 950, "color": "FFDDD5"},
        {"text": "Engaged audience > offers", "size": 950, "color": "FFDDD5"},
    ])
    return wrap_slide(shapes)
