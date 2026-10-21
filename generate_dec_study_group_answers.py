#!/usr/bin/env python3
"""
generate_dec_study_group_answers.py
Generates comprehensive, publication-grade Study Group Answers for:
STUDY GROUP QUESTIONS (OCTOBER 19TH – 25TH 2026)
CELL LEADERS TRAINING DECEMBER 2021 – FAITHFULNESS IN MINISTRY.

Adheres strictly to the KeyIndicators KPI Standard and Saints Community Church
Study Group Guidelines:
- Header Metadata (Name: AGBEDE STEPHEN AYOTOMIWA, Church: SAINTS COMMUNITY CHURCH, AKURE)
- Four-part architecture: Verbatim Question -> Introduction -> Multi-tiered Body -> Conclusion
- Greek/Hebrew exegetical precision (pistos, pisteuo, eritheia, dokimazo, dokimos, adokimos, hypomone, typos, episkope)
- Indented Scripture blockquotes with bold-italic references and verse numbers
- Letter page geometry, Arial 12pt typography, dark neutral palette
"""

import os
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def create_study_answers(docx_paths, md_paths):
    doc = Document()

    # Page setup: Standard Letter (8.5" x 11.0"), 1.0" margins
    for section in doc.sections:
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Base style
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)
    font.color.rgb = RGBColor(0x11, 0x18, 0x27) # Dark neutral

    md_lines = []

    def add_title(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(18)
        md_lines.append(f"# {text}")

    def add_subtitle(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(8)
        run = p.add_run(text)
        run.bold = True
        run.italic = True
        run.font.name = 'Arial'
        run.font.size = Pt(14)
        md_lines.append(f"## *{text}*\n")

    def add_metadata(name, church):
        p1 = doc.add_paragraph()
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(2)
        r1 = p1.add_run(f"Name: {name}")
        r1.bold = True
        r1.font.name = 'Arial'
        r1.font.size = Pt(11)

        p2 = doc.add_paragraph()
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(12)
        r2 = p2.add_run(f"CHURCH: {church}")
        r2.bold = True
        r2.font.name = 'Arial'
        r2.font.size = Pt(11)

        md_lines.append(f"**Name:** {name}  ")
        md_lines.append(f"**CHURCH:** {church}  \n\n---\n")

    def add_q_header(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(14)
        md_lines.append(f"## {text}\n")
        return p

    def add_q_prompt(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(12)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"**{text}**\n")
        return p

    def add_section_h(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(13)
        md_lines.append(f"### {text}\n")
        return p

    def add_body_subh(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"#### {text}\n")
        return p

    def ref(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(text)
        run.bold = True
        run.italic = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"***{text}***")
        return p

    def v(text):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        run = p.add_run(text)
        run.bold = True
        run.italic = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"> ***{text}***  ")
        return p

    def sp():
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        md_lines.append("")
        return p

    def b(runs_data):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        md_para = ""
        for r_text, r_bold, r_italic in runs_data:
            run = p.add_run(r_text)
            run.font.name = 'Arial'
            run.font.size = Pt(12)
            if r_bold:
                run.bold = True
            if r_italic:
                run.italic = True

            seg = r_text
            if r_bold and r_italic:
                seg = f"***{seg}***"
            elif r_bold:
                seg = f"**{seg}**"
            elif r_italic:
                seg = f"*{seg}*"
            md_para += seg

        md_lines.append(md_para + "\n")
        return p

    # =========================================================================
    # DOCUMENT HEADER
    # =========================================================================
    add_title("STUDY GROUP QUESTIONS (OCTOBER 19TH – 25TH 2026)")
    add_subtitle("CELL LEADERS TRAINING DECEMBER 2021 – FAITHFULNESS IN MINISTRY")
    add_metadata("AGBEDE STEPHEN AYOTOMIWA", "SAINTS COMMUNITY CHURCH, AKURE")

    # =========================================================================
    # QUESTION 1
    # =========================================================================
    add_q_header("QUESTION 1")
    add_q_prompt("Explain the difference between faithfulness and ambition using the examples of Ruth and Naomi, Elisha and Elijah, Joshua and Moses, Timoth and Paul as done within this track.")

    add_section_h("INTRODUCTION")
    b([
        ("Within the curriculum of the local church and the systematic discipling of Christian workers, no subject demands greater exegetical precision and heart examination than the spiritual nature of ministerial motivation. In the contemporary church climate, human zeal, charismatic assertiveness, and the desire for ecclesiastical visibility are frequently mistaken for spiritual devotion. Believers and emerging leaders often embark on ministerial service with mixed motives, projecting an image of intense dedication while subtly harboring self-centered ambitions. The teaching ", False, False),
        ("Cell Leaders Training: Faithfulness in Ministry", True, False),
        (" cuts through this religious ambiguity by drawing an absolute, unyielding demarcation between ", False, False),
        ("faithfulness", True, False),
        (" and ", False, False),
        ("ambition", True, False),
        (". Scriptural faithfulness is not passive acquiescence, nor is it a temporary display of enthusiasm designed to secure personal promotion. Faithfulness is unwavering loyalty to a divine cause, anchored in pure servanthood, regardless of whether personal desires are fulfilled or visibility is attained. Ambition, conversely, is transactional, self-advancing, and conditional—altering its posture the moment personal expectations are denied. To thoroughly expound this distinction, this analysis establishes the doctrinal and lexical foundation of stewardship from 1 Corinthians 4:1-2, unpacks the four biblical narratives employed in the teaching (Ruth and Naomi, Elisha and Elijah, Joshua and Moses, and Timothy and Paul), and articulates the direct ministerial application required of every cell leader and worker in the local assembly.", False, False)
    ])
    sp()

    add_section_h("BODY")

    add_body_subh("Foundational Definitions: Faithfulness (*Pistos*) vs. Carnal Ambition (*Eritheia*)")
    ref("1 Corinthians 4:1-2")
    v("[1] Let a man so account of us, as of the ministers of Christ, and stewards of the mysteries of God.")
    v("[2] Moreover it is required in stewards, that a man be found faithful.")
    sp()

    b([
        ("The Apostle Paul establishes the governing canon of ministerial evaluation in ", False, False),
        ("1 Corinthians 4:1-2", True, False),
        (". Ministers are not evaluated by worldly metrics of crowd size, oratorical charisma, or ecclesiastical title, but as servants of Christ and stewards (", False, False),
        ("oikonomos", True, True),
        (") of divine mysteries. The single non-negotiable requirement laid upon a steward is that they be found ", False, False),
        ("faithful", True, False),
        (". In the Greek text, 'faithful' is the adjective ", False, False),
        ("pistos", True, True),
        (", fundamentally linked to the verb ", False, False),
        ("pisteuo", True, True),
        (" (to believe, trust, commit). ", False, False),
        ("Pistos", True, True),
        (" denotes someone who is utterly reliable, steadfast, trustworthy, and constant under every condition. In 1 Corinthians 10:13, Scripture declares that God is faithful (", False, False),
        ("pistos de ho Theos", True, True),
        (")—meaning God never alters His commitment, never reneges on His covenant, and remains perpetually dependable. When applied to human stewards, faithfulness is synonymous with ", False, False),
        ("absolute loyalty", True, False),
        (": a one-hundred-percent dedication to God, the assignment, the local church, and spiritual authority, irrespective of hardships or changing circumstances.", False, False)
    ])
    sp()

    b([
        ("In sharp contrast, carnal ambition—described in the New Testament as ", False, False),
        ("eritheia", True, True),
        (" (self-seeking rivalry, factional ambition, partisan spirit, as in Philippians 2:3 and James 3:14-16)—is inherently self-absorbed. The ambitious person enters ministry with a covert agenda: they seek to rise through the ranks, gain proximity to senior leadership, acquire a platform, or carve out an independent reputation. The teaching provides the definitive spiritual diagnostic to unmask ambition: ", False, False),
        ("the true nature of an individual's service is revealed when their personal desires, preferences, and positional ambitions are denied or delayed", True, False),
        (". An ambitious worker will act faithful only as long as the trajectory favors their advancement; the moment an alteration occurs, a correction is administered, or another person is preferred, their disposition changes immediately. They murmur, withdraw their service, exhibit resentment, or orchestrate rebellion. The faithful steward, however, remains immovable because their service was never tethered to personal entitlement.", False, False)
    ])
    sp()

    add_body_subh("Case Study 1: Ruth and Naomi — Uncalculating Loyalty vs. Conditional Affection")
    ref("Ruth 1:14, 16-17")
    v("[14] And they lifted up their voice, and wept again: and Orpah kissed her mother in law; but Ruth clave unto her.")
    v("[16] And Ruth said, Intreat me not to leave thee, or to return from following after thee: for whither thou goest, I will go; and where thou lodgest, I will lodge: thy people shall be my people, and thy God my God:")
    v("[17] Where thou diest, will I die, and there will I be buried: the LORD do so to me, and more also, if ought but death part thee and me.")
    sp()

    b([
        ("The narrative of Ruth and Naomi provides an unassailable demonstration of pure faithfulness contrasted with conditional attachment. Following the death of Elimelech and his two sons in Moab, Naomi stood destitute, bereft of husband, sons, and financial security. When Naomi urged her two Moabite daughters-in-law to return to their maternal homes where prospective marriage and societal security awaited them, Orpah kissed Naomi and departed, but Ruth clave unto her (", False, False),
        ("Ruth 1:14", True, False),
        ("). Orpah represents affection that terminates when personal advantage ceases. Orpah loved Naomi, but her commitment was contingent upon what the relationship could provide. When Naomi could offer no more husbands, inheritance, or social elevation, Orpah's devotion evaporated.", False, False)
    ])
    sp()

    b([
        ("Ruth, however, epitomizes pure, uncalculating faithfulness. In verse 16-17, she articulated an unconditional covenant that bound her life, geography, culture, spiritual allegiance, and mortality to Naomi. Ruth was a Gentile widow entering an Israelite society that historically despised Moabites; she possessed zero earthly guarantee of marriage, property, or recognition. Yet she declared: ", False, False),
        ("*'Whither thou goest, I will go; and where thou lodgest, I will lodge: thy people shall be my people, and thy God my God.'*", True, True),
        (" Ruth was not calculating perks, calculating an inheritance, or pursuing upward social mobility. Her service was unadulterated loyalty. Ambition asks: 'What will I gain from this association?' Faithfulness declares: 'I am committed to this spiritual assignment regardless of personal cost.'", False, False)
    ])
    sp()

    add_body_subh("Case Study 2: Elisha and Elijah — Humble Servanthood vs. Independent Platforms")
    ref("2 Kings 3:11")
    v("[11] But Jehoshaphat said, Is there not here a prophet of the LORD, that we may enquire of the LORD by him? And one of the king of Israel's servants answered and said, Here is Elisha the son of Shaphat, which poured water on the hands of Elijah.")
    sp()

    b([
        ("The relationship between Elisha and Elijah reveals the structural posture of faithful discipleship. When Elijah cast his mantle upon Elisha in 1 Kings 19:19-21, Elisha was plowing with twelve yoke of oxen—indicating that he hailed from a wealthy, established agricultural family. Elisha slew his oxen, burned the farming implements to feed the people, and departed to minister unto Elijah. He surrendered familial wealth and social independence to assume the position of a domestic attendant. When the kings of Israel, Judah, and Edom sought a genuine word from the Lord years later in 2 Kings 3:11, the identifying credential that validated Elisha was not oratorical charisma or prophetic self-promotion, but: ", False, False),
        ("*'Here is Elisha the son of Shaphat, which poured water on the hands of Elijah.'*", True, True),
        (" He was known across the nation as the faithful servant who washed the hands of his spiritual father.", False, False)
    ])
    sp()

    ref("2 Kings 2:2, 14")
    v("[2] And Elijah said unto Elisha, Tarry here, I pray thee; for the LORD hath sent me to Bethel. And Elisha said unto him, As the LORD liveth, and as thy soul liveth, I will not leave thee. So they went down to Bethel.")
    v("[14] And he took the mantle of Elijah that fell from him, and smote the waters, and said, Where is the LORD God of Elijah? and when he also had smitten the waters, they parted hither and thither: and Elisha went over.")
    sp()

    b([
        ("The crucial divergence between faithfulness and ambition is showcased in 2 Kings 2 as Elijah journeyed from Gilgal to Bethel, Jericho, and the Jordan. At each juncture, the 'sons of the prophets' stood at a distance, taunting Elisha with cynical commentary: 'Knowest thou that the LORD will take away thy master from thy head today?' The sons of the prophets represent religious onlookers—individuals who possess theoretical prophetic information but lack relational loyalty and sacrificial endurance. They remained spectators on the bank. Elisha, however, refused to abandon his duty post, answering repeatedly: ", False, False),
        ("*'As the LORD liveth, and as thy soul liveth, I will not leave thee.'*", True, True),
        (" When Elijah asked what he should do for him before his departure, Elisha requested a double portion of his spirit. The teaching makes a vital doctrinal clarification: ", False, False),
        ("this was not an ambitious demand to outshine or overshadow Elijah", True, False),
        ("; under Mosaic law, the double portion was the exclusive legal right of the firstborn son, signifying Elisha's desire to inherit the spiritual responsibility of maintaining the continuity of Elijah's prophetic office.", False, False)
    ])
    sp()

    b([
        ("The crowning proof of Elisha's pure motivation occurred immediately after Elijah was taken up. Picking up the fallen mantle, Elisha struck the Jordan and cried: ", False, False),
        ("*'Where is the LORD God of Elijah?'*", True, True),
        (" He did not cry: 'Where is the God of Elisha? Now my time has come! Finally, the old man is gone and my personal ministry can shine!' Elisha wore Elijah's ministry, honored Elijah's legacy, and pointed the people back to the God of his spiritual father. Ambition seeks an independent brand; faithfulness honors spiritual lineage and builds upon the apostolic foundation.", False, False)
    ])
    sp()

    add_body_subh("Case Study 3: Joshua and Moses — Generational Honor vs. Revisionist Usurpation")
    ref("Joshua 1:1-2")
    v("[1] Now after the death of Moses the servant of the LORD it came to pass, that the LORD spake unto Joshua the son of Nun, Moses' minister, saying,")
    v("[2] Moses my servant is dead; now therefore arise, go over this Jordan, thou, and all this people, unto the land which I do give to them, even to the children of Israel.")
    sp()

    b([
        ("A parallel dynamic of generational fidelity is demonstrated in Joshua's lifelong walk with Moses. Throughout Exodus and Numbers, Joshua was designated not as a co-leader or political rival, but as ", False, False),
        ("Moses' minister", True, False),
        (" (", False, False),
        ("Exodus 24:13; 33:11", True, False),
        (")—a dedicated assistant who remained in the tabernacle, guarded Moses' integrity, and executed military commands at Rephidim. When Moses passed away on Mount Nebo, God called Joshua to lead Israel across the Jordan (", False, False),
        ("Joshua 1:1-2", True, False),
        ("). An ambitious subordinate, upon ascending to supreme leadership, frequently engages in revisionist history—minimizing the predecessor's accomplishments, magnifying their flaws, and presenting themselves as the enlightened emancipator. An ambitious successor would have proclaimed to Israel: 'Moses failed to bring you into the Promised Land because of his anger at Meribah; but I am the new anointed vessel who will accomplish what Moses could never do!'", False, False)
    ])
    sp()

    b([
        ("Joshua completely rejected such fleshly self-aggrandizement. He anchored his leadership strictly in the continuity of Moses' divine instructions. Throughout the conquest of Canaan, Joshua repeatedly affirmed that the Lord God was the God of Moses, and that Israel was executing the precise commandments delivered through Moses the servant of the Lord (", False, False),
        ("Joshua 1:7; 4:10; 11:15", True, False),
        ("). The teaching notes: ", False, False),
        ("loyal, faithful people do not do otherwise when the leader is no longer present", True, False),
        (". Joshua did not rewrite the doctrine, alter the pattern, or demean the father of the work; he faithfully stewarded the vision across Jordan. Ambition uses the departure of a leader to establish personal supremacy; faithfulness guards the deposit and preserves apostolic continuity.", False, False)
    ])
    sp()

    add_body_subh("Case Study 4: Timothy and Paul — Sonship and Selfless Care vs. Mercenary Self-Seeking")
    ref("Philippians 2:19-22")
    v("[19] But I trust in the Lord Jesus to send Timotheus shortly unto you, that I also may be of good comfort, when I know your state.")
    v("[20] For I have no man likeminded, who will naturally care for your state.")
    v("[21] For all seek their own, not the things which are Jesus Christ's.")
    v("[22] But ye know the proof of him, that, as a son with the father, he hath served with me in the gospel.")
    sp()

    b([
        ("The New Testament archetype of faithful ministerial partnership is found in Paul's commendation of Timothy. In ", False, False),
        ("Philippians 2:19-22", True, False),
        (", Paul draws an explicit contrast between the pervasive culture of carnal ambition and Timothy's selfless faithfulness. Paul lamented the general ministerial climate of his day, delivering a staggering diagnosis in verse 21: ", False, False),
        ("*'For all seek their own, not the things which are Jesus Christ's.'*", True, True),
        (" Many ministers in the first century were engaged in ministry as a commercial career or personal platform; they sought their own reputation, financial comfort, and status. But concerning Timothy, Paul declared: ", False, False),
        ("*'For I have no man likeminded, who will naturally care for your state... But ye know the proof of him, that, as a son with the father, he hath served with me in the gospel.'*", True, True),
        (" The Greek word for 'likeminded' is ", False, False),
        ("isopsychos", True, True),
        ("—sharing the exact same soul, heartbeat, and ministerial priorities.", False, False)
    ])
    sp()

    b([
        ("The teaching observes that Timothy was so thoroughly integrated into Paul's apostolic labor that one could readily refer to Paul's epistles as 'Timothy's epistles.' In 1 and 2 Thessalonians, 2 Corinthians, Philippians, Colossians, and Philemon, Timothy's name is conjoined with Paul in the opening salutations. Timothy was sent into troubled churches (Corinth, Thessalonica, Philippi, Ephesus) not to build an independent brand, but to remind believers of Paul's ways which be in Christ (", False, False),
        ("1 Corinthians 4:17", True, False),
        ("). Timothy served ", False, False),
        ("as a son with a father", True, False),
        ("—a posture of covenantal honor, absolute transparency, and shared labor. Ambition operates as a mercenary: it works for wages, demands recognition, and jumps to the highest bidder. Faithfulness operates as a son: it invests life, protects the house, and naturally cares for the sheep.", False, False)
    ])
    sp()

    add_body_subh("Direct Ministerial & Discipleship Application: The Behavior of Faithful Stewards")
    b([
        ("From these four case studies, the teaching deduces the practical, everyday conduct that separates the faithful cell leader from the ambitious worker:", False, False)
    ])
    sp()

    b([
        ("1. The Absence Test: ", True, False),
        ("The true mark of a faithful worker is what they do when senior leadership is completely absent. The teaching declares: ", False, False),
        ("*'Faithful people will do in the service what they would have done if Jesus was seated in the front seat. Faithful people will do in the service what they would have done if their pastor was seated in the service. They do things behind you as though you are seated there.'*", True, True),
        (" The ambitious worker is performative: they arrive early, pray loudly, and minister intensely only when the pastor or coordinator is watching, but lapse into indolence, gossip, or flippancy when unobserved. The faithful leader works under the perpetual gaze of God.", False, False)
    ])
    sp()

    b([
        ("2. The Promotion and Alteration Test: ", True, False),
        ("When an ambitious person is bypassed for promotion, assigned to a menial department (such as sanitation, security, or operations), or corrected by leadership, they take offense, cause division, and eventually leave. The faithful steward rejoices in any opportunity to serve the body of Christ, knowing that no task in the house of God is beneath a servant of Jesus.", False, False)
    ])
    sp()

    b([
        ("3. The Flock Protection Test: ", True, False),
        ("The ambitious leader treats church members as personal property, using them to validate their leadership prowess or poaching them from other assemblies. The faithful steward wins souls from scratch (Romans 15:20), nurtures them with patience, and faithfully integrates them into the local church family.", False, False)
    ])
    sp()

    add_section_h("CONCLUSION")
    b([
        ("In conclusion, faithfulness and ambition represent two diametrically opposed spirits within the work of the ministry. Ambition originates in the carnal ego (", False, False),
        ("eritheia", True, True),
        ("); it is transactional, self-seeking, impatient of process, and easily offended when personal desires are frustrated. It leaves behind a trail of division, broken relationships, and spiritual sterility—epitomized by the sons of the prophets who stood idly by while the chariot of fire descended. Faithfulness originates in the character of God (", False, False),
        ("pistos", True, True),
        ("); it is selfless, patient, resilient under pressure, and steadfastly loyal across generations—demonstrated in Ruth's covenantal devotion, Elisha's humble service, Joshua's generational honor, and Timothy's filial dedication.", False, False)
    ])
    sp()

    b([
        ("Every cell leader and worker-in-training must decisively crucify personal ambition on the altar of sacrificial stewardship. We must settle in our hearts that our names, reputations, and comfort are of no consequence; the only reality that matters is that we be found faithful stewards of the mysteries of God. Let our service be rendered with such unblemished integrity that whether observed by men or laboring in total obscurity, our posture remains constant, our loyalty unwavering, and our devotion wholly dedicated to the Lord Jesus Christ and His church.", False, False)
    ])
    sp()

    # =========================================================================
    # QUESTION 2
    # =========================================================================
    add_q_header("QUESTION 2")
    add_q_prompt("Does God test believers? With the aid of this track explain the correlation between tests and faithfulness in the work of the ministry. Also explain the role of tests in attaining leadership. Using at least 2 biblical characters employed in this teaching, explain the possible outcomes to tests.")

    add_section_h("INTRODUCTION")
    b([
        ("A pervasive doctrinal error within contemporary charismatic Christianity is the unscriptural presumption that under the New Covenant of grace, the believer is entirely exempt from divine examination, testing, or trials. Many confuse the clear scriptural truth that God never tempts believers with moral evil (James 1:13) with the false assumption that God never proves or tests their character. Consequently, when adversity, ministerial inconvenience, delayed recognition, or intense pressure confront a worker, they mistakenly attribute the friction solely to demonic attack or conclude that God has abandoned them. The teaching ", False, False),
        ("Cell Leaders Training: Faithfulness in Ministry", True, False),
        (" provides a rigorous theological correction to this fallacy. God does indeed test believers. Divine testing is not an adversarial scheme to orchestrate destruction; rather, it is a deliberate, necessary pedagogical process designed to prove structural integrity, forge spiritual stamina, and validate character before authority is conferred. Faithfulness cannot exist as a theoretical sentiment; it can only be demonstrated and authenticated through the crucible of testing. To provide an exhaustive answer, this exposition establishes the doctrinal reality of divine testing, traces the direct correlation between testing and ministerial faithfulness, unpacks the mandatory role of vetting in attaining leadership, and contrasts the divergent outcomes of testing through the lives of Joseph, Gehazi, and David.", False, False)
    ])
    sp()

    add_section_h("BODY")

    add_body_subh("Does God Test Believers? The Biblical Theology of Divine Proving (*Dokimazo*)")
    ref("James 1:2-4")
    v("[2] My brethren, count it all joy when ye fall into divers temptations;")
    v("[3] Knowing this, that the trying of your faith worketh patience.")
    v("[4] But let patience have her perfect work, that ye may be perfect and entire, wanting nothing.")
    sp()

    b([
        ("The question 'Does God test believers?' must be answered with an unqualified ", False, False),
        ("YES", True, False),
        (", provided the nature of divine testing is accurately defined according to Scripture. The Bible distinguishes sharply between two distinct concepts: ", False, False),
        ("solicitation to evil", True, False),
        (" (which God never does) and ", False, False),
        ("examination of character", True, False),
        (" (which God consistently executes). In ", False, False),
        ("James 1:13", True, False),
        (", Scripture affirms: ", False, False),
        ("*'Let no man say when he is tempted, I am tempted of God: for God cannot be tempted with evil, neither tempteth he any man.'*", True, True),
        (" God never orchestrates moral sin, sickness, or destruction. However, in the very same chapter (", False, False),
        ("James 1:2-3", True, False),
        ("), James commands believers to count it all joy when they encounter diverse trials (", False, False),
        ("peirasmos", True, True),
        ("), knowing that the trying (", False, False),
        ("dokimion", True, True),
        (") of their faith produces steadfast patience (", False, False),
        ("hypomone", True, True),
        (").", False, False)
    ])
    sp()

    b([
        ("The Greek word family for biblical testing—", False, False),
        ("dokimazo", True, True),
        (" (verb), ", False, False),
        ("dokime", True, True),
        (" (noun), and ", False, False),
        ("dokimos", True, True),
        (" (adjective)—originates from the metallurgy of the ancient world. It refers to subjecting gold or silver to extreme furnace heat to burn away dross and verify that the metal is authentic and meets the legal standard. In secular and commercial usage, it was applied to testing coins to ensure they were not counterfeit. Therefore, when God tests a believer, ", False, False),
        ("the purpose is never to make them fail; the purpose is to prove their authenticity, strengthen their spiritual fiber, and display their integrity", True, False),
        (". The teaching employs a vivid contemporary analogy: ", False, False),
        ("*'God tests people not so that they can fail, but the way an automobile engineer test-drives a newly manufactured vehicle.'*", True, True),
        (" The car is driven over rugged terrain, tested at high speeds, and subjected to emergency braking not to destroy the engine, but to prove that it can handle real-world conditions without breaking down. In Job 1:8, God boasted of Job's moral fortitude before Satan: ", False, False),
        ("*'Hast thou considered my servant Job, that there is none like him in the earth, a perfect and an upright man?'*", True, True),
        (" God knew Job's integrity would survive the storm.", False, False)
    ])
    sp()

    ref("1 Samuel 16:6-7")
    v("[6] And it came to pass, when they were come, that he looked on Eliab, and said, Surely the LORD'S anointed is before him.")
    v("[7] But the LORD said unto Samuel, Look not on his countenance, or on the height of his stature; because I have refused him: for the LORD seeth not as man seeth; for man looketh on the outward appearance, but the LORD looketh on the heart.")
    sp()

    b([
        ("The Old Testament confirms this principle in the rejection of Eliab in ", False, False),
        ("1 Samuel 16:6-7", True, False),
        (". When Samuel arrived in Bethlehem, he was captivated by Eliab's regal height and handsome countenance, declaring: 'Surely the LORD's anointed is before him!' But the Lord said: ", False, False),
        ("*'Look not on his countenance, or on the height of his stature; because I have refused him.'*", True, True),
        (" The Hebrew word for 'refused' is ", False, False),
        ("ma'as", True, True),
        (", signifying that Eliab had been examined and rejected as substandard. Eliab looked like a king on the exterior, but within his heart lay anger, pride, and contempt (as later manifested in his vicious attack on David in 1 Samuel 17:28). Eliab failed the internal character test. God tests the heart because true spiritual ministry cannot be built upon cosmetic appearance.", False, False)
    ])
    sp()

    add_body_subh("The Correlation Between Tests and Faithfulness in the Work of Ministry")
    b([
        ("Faithfulness (", False, False),
        ("pistos", True, True),
        (") and testing (", False, False),
        ("dokimazo", True, True),
        (") exist in an unbreakable symbiotic relationship. One cannot speak of faithfulness in the abstract. An untested soldier cannot claim courage; an untested bridge cannot guarantee structural safety; and an untested believer cannot boast of faithfulness. Faithfulness is only demonstrated when opposing forces exert pressure upon a leader's conviction.", False, False)
    ])
    sp()

    b([
        ("The teaching outlines four key dimensions where tests correlate directly with ministerial faithfulness:", False, False)
    ])
    sp()

    b([
        ("1. Tests Separate Sentiment from Conviction: ", True, False),
        ("Anyone can sing songs of consecration, shout 'amen' in a convention, or profess undying loyalty when the sun is shining and blessings abound. But faithfulness is only verified when adversity strikes—when funds are low, when physical fatigue sets in, when disciples are ungrateful, or when personal sacrifice is demanded. If a leader abandons their cell group because of rain, minor headaches, or internet data constraints, they have failed the test of faithfulness.", False, False)
    ])
    sp()

    b([
        ("2. Tests Develop Spiritual Endurance (*Hypomone*): ", True, False),
        ("Ministry is not a hundred-meter sprint; it is an exhausting, lifelong marathon. James 1:3-4 declares that the trying of faith produces patience (", False, False),
        ("hypomone", True, True),
        (")—the capacity to stand under heavy burdens without buckling. When a leader runs away from tests, conflicts, or difficult people, they run away from the very instrument God uses to build spiritual muscles. Untested leaders buckle at the first sign of pastoral crisis.", False, False)
    ])
    sp()

    b([
        ("3. Tests Guard Against Pride: ", True, False),
        ("Unchecked success breeds arrogance. Testing reminds the minister of their human frailty, forcing them to depend wholly upon the Holy Spirit and the Word rather than their natural intelligence or personality. Paul's thorn in the flesh (2 Corinthians 12:7-10) was permitted so that he would not be exalted above measure, discovering that God's grace is perfected in human weakness.", False, False)
    ])
    sp()

    b([
        ("4. Tests Qualify the Leader for Multiplication: ", True, False),
        ("Jesus declared in Luke 16:10: ", False, False),
        ("*'He that is faithful in that which is least is faithful also in much.'*", True, True),
        (" God does not entrust fifty or a hundred souls to a leader who starves or mismanages five. The test of the 'few' is the mandatory qualification for the stewardship of the 'many.'", False, False)
    ])
    sp()

    add_body_subh("The Role of Tests in Attaining Leadership: The Mandate of Vetting (*Dokimazo*)")
    ref("1 Timothy 3:10, 13")
    v("[10] And let these also first be proved; then let them use the office of a deacon, being found blameless.")
    v("[13] For they that have used the office of a deacon well purchase to themselves a good degree, and great boldness in the faith which is in Christ Jesus.")
    sp()

    b([
        ("In the administrative government of the local church, testing is not an optional luxury; it is a mandatory apostolic statute. In ", False, False),
        ("1 Timothy 3:10", True, False),
        (", Paul issues an explicit command concerning prospective deacons and ministry associates: ", False, False),
        ("let these also first be proved", True, False),
        (" (", False, False),
        ("dokimazesthosan proton", True, True),
        ("). Notice the chronological priority: ", False, False),
        ("proving must precede appointment", True, False),
        (". The church must never donate responsibility to unproven individuals simply because they are wealthy, academically educated, socially influential, or loudly enthusiastic.", False, False)
    ])
    sp()

    b([
        ("The teaching asks a penetrating question: ", False, False),
        ("*'When Paul writes that deacons must first be proved, who does the proving? Does the Holy Spirit send an angel with a certificate? No! Human leadership must do the proving!'*", True, True),
        (" Pastors and leaders must actively observe prospective workers in real-life settings: How do they handle menial tasks? Are they punctual? How do they react to correction? Do they finish what they start? Can they handle small budgets without dishonesty? In the local church, genuine leadership development starts at the lowest levels—sanitation, ushering, works, operations, logistics. When title-seeking individuals join an assembly seeking rapid ordination and are handed a broom or asked to clean restrooms, their true heart is exposed immediately. If they flee, the church is spared future calamity; if they serve with joy, they purchase to themselves a good degree and great boldness in the faith (", False, False),
        ("1 Timothy 3:13", True, False),
        (").", False, False)
    ])
    sp()

    ref("2 Timothy 2:15")
    v("[15] Study to shew thyself approved unto God, a workman that needeth not to be ashamed, rightly dividing the word of truth.")
    sp()

    b([
        ("In ", False, False),
        ("2 Timothy 2:15", True, False),
        (", Paul reiterates that the goal of every minister is to be ", False, False),
        ("approved unto God", True, False),
        (". The Greek word is ", False, False),
        ("dokimos", True, True),
        ("—standing approved after being thoroughly tested in the fire. In Acts 15:26, when the Jerusalem council sent leaders to Antioch, they commended Paul and Barnabas as ", False, False),
        ("*'men that have hazarded their lives for the name of our Lord Jesus Christ.'*", True, True),
        (" Their authority was recognized because they had been proven under persecution. The teaching delivers a stern pastoral warning: ", False, False),
        ("pastors must never bypass the testing process", True, False),
        (". God allows fires and wilderness experiences so that those who cannot survive the desert do not cross over into the Promised Land to corrupt the flock.", False, False)
    ])
    sp()

    add_body_subh("Biblical Characters and Possible Outcomes to Tests: Joseph vs. Gehazi")
    b([
        ("The teaching examines biblical characters to demonstrate the two diametrically opposed outcomes of divine testing: passing into exaltation and spiritual maturity (", False, False),
        ("dokimos", True, True),
        (") versus failing into compromise, greed, and spiritual ruin (", False, False),
        ("adokimos", True, True),
        (").", False, False)
    ])
    sp()

    b([
        ("Outcome 1: The Passing of Tests — Joseph (Endurance, Integrity, and Exaltation):", True, False)
    ])
    sp()

    b([
        ("Joseph stands as the supreme Old Testament archetype of a leader who passed consecutive, escalating tests across diverse environments:", False, False)
    ])
    sp()

    b([
        ("• The Test of Servitude in Potiphar's House: ", True, False),
        ("Sold into Egyptian slavery by his brothers, Joseph did not collapse into depression, bitterness, or rebellion. He served with such excellence and reliability that Potiphar left everything in his hands. When Potiphar's wife repeatedly solicited him to commit adultery, Joseph passed the moral test of secret integrity, declaring: ", False, False),
        ("*'How then can I do this great wickedness, and sin against God?' (Genesis 39:9)*", True, True),
        (". He chose prison over moral compromise.", False, False)
    ])
    sp()

    b([
        ("• The Test of Injustice and Obscurity in the Dungeon: ", True, False),
        ("Falsely accused and incarcerated, Joseph did not abandon his ministry. He did not say: 'Is this the reward for serving God? Dreams got me into trouble; I will never interpret another dream!' Instead, he observed the distress of the baker and butler, ministered to their needs with compassion, and accurately interpreted their dreams. Even when the butler forgot him for two full years (Genesis 40:23), Joseph nurtured no bitterness.", False, False)
    ])
    sp()

    b([
        ("• The Test of Power and Vindication before Pharaoh and his Brothers: ", True, False),
        ("When summoned before Pharaoh, Joseph did not bargain for his release or demand retribution against Potiphar's wife; he pointed Pharaoh directly to God (Genesis 41:16). When exalted to govern the empire, he fed Egypt, provided for Potiphar's household, and when his fearful brothers stood before him, Joseph revealed the supreme character of Christ: ", False, False),
        ("*'Fear not: for am I in the place of God? But as for you, ye thought evil against me; but God meant it unto good, to save much people alive' (Genesis 50:19-20)*", True, True),
        (". Because Joseph was faithful across every station, God entrusted him with the salvation of nations. Tested faithfulness produces unshakeable character and kingdom exaltation.", False, False)
    ])
    sp()

    b([
        ("Outcome 2: The Failure of Tests — Gehazi (Covetousness, Deceit, and Disqualification):", True, False)
    ])
    sp()

    ref("2 Kings 5:20, 26-27")
    v("[20] But Gehazi, the servant of Elisha the man of God, said, Behold, my master hath spared Naaman this Syrian, in not receiving at his hands that which he brought: but, as the LORD liveth, I will run after him, and take somewhat of him.")
    v("[26] And he said unto him, Went not mine heart with thee, when the man turned again from his chariot to meet thee? Is it a time to receive money, and to receive garments, and oliveyards, and vineyards, and sheep, and oxen, and menservants, and maidservants?")
    v("[27] The leprosy therefore of Naaman shall cleave unto thee, and unto thy seed for ever. And he went out from his presence a leper as white as snow.")
    sp()

    b([
        ("In tragic contrast, Gehazi exemplifies the catastrophic failure of a minister under testing. Gehazi occupied an extraordinary spiritual position as the personal associate and disciple of Elisha. He witnessed dead children raised to life, poisoned pots cleansed, and miracles multiplied; he stood in prime position to inherit the spiritual mantle and double portion from Elisha, just as Elisha had inherited from Elijah. But Gehazi was untested in his heart concerning material wealth.", False, False)
    ])
    sp()

    b([
        ("When Naaman the Syrian commander was miraculously cured of leprosy, Elisha adamantly refused to accept monetary remuneration, ensuring that the grace of Yahweh was seen as entirely free and unmerited. But Gehazi failed the test of covetousness. In ", False, False),
        ("2 Kings 5:20", True, False),
        (", Gehazi swore a carnal oath: ", False, False),
        ("*'As the LORD liveth, I will run after him, and take somewhat of him.'*", True, True),
        (" Gehazi pursued Naaman, fabricated a deceitful story about two young prophets arriving from Mount Ephraim, and collected two talents of silver and two changes of garments, hiding them secretly in his house. Elisha confronted him in the Spirit: ", False, False),
        ("*'Went not mine heart with thee, when the man turned again from his chariot to meet thee? Is it a time to receive money, and to receive garments... The leprosy therefore of Naaman shall cleave unto thee, and unto thy seed for ever.'*", True, True),
        (" Gehazi walked out of Elisha's presence a leper as white as snow.", False, False)
    ])
    sp()

    b([
        ("The lesson of Gehazi is chilling for every Christian worker: ", False, False),
        ("proximity to an anointed ministry does not guarantee spiritual integrity", True, False),
        (". An untested heart, harbored covetousness, and the desire to exploit the work of God for personal financial gain will disqualify a leader and bring spiritual leprosy upon their lineage. Gehazi became ", False, False),
        ("adokimos", True, True),
        ("—disqualified, rejected, and ruined.", False, False)
    ])
    sp()

    add_body_subh("Supplementary Character: David and the Defense of the Few Sheep")
    ref("1 Samuel 17:34-36")
    v("[34] And David said unto Saul, Thy servant kept his father's sheep, and there came a lion, and a bear, and took a lamb out of the flock:")
    v("[35] And I went out after him, and smote him, and delivered it out of his mouth: and when he arose against me, I caught him by his beard, and smote him, and slew him.")
    v("[36] Thy servant slew both the lion and the bear: and this uncircumcised Philistine shall be as one of them.")
    sp()

    b([
        ("David provides an additional illustration of passing tests in obscurity. Long before David confronted Goliath on the national stage, he was tested in the wilderness with his father's 'few sheep.' When a lion and a bear attacked the flock, David did not abandon the sheep to save himself; he pursued the beasts, delivered the lamb, and slew them. He passed the test of pastoral courage, vigilance, and sacrificial love when no human audience was watching. Because he was faithful in the wilderness test, God promoted him to slay Goliath and shepherd the nation of Israel. Private tests successfully navigated lead to public spiritual authority.", False, False)
    ])
    sp()

    add_section_h("CONCLUSION")
    b([
        ("In summary, God actively tests believers not to orchestrate their downfall, but to prove, refine, and authenticate their character for spiritual stewardship. The correlation between tests and faithfulness is absolute: faithfulness cannot be verified in ease; it is forged and demonstrated solely through the pressure of trials. In the government of the local church, testing (", False, False),
        ("dokimazo", True, True),
        (") is the indispensable prerequisite for leadership appointment. Untested leadership is a hazard to the flock.", False, False)
    ])
    sp()

    b([
        ("The biblical narratives place before every cell leader and worker two distinct outcomes. We can follow the path of Joseph and David—enduring trials without bitterness, maintaining moral purity in obscurity, protecting the few sheep, and emerging approved (", False, False),
        ("dokimos", True, True),
        (") to shepherd nations. Or we can succumb to the path of Eliab and Gehazi—harboring secret pride, coveting material gain, failing the test of integrity, and ending in disqualification (", False, False),
        ("adokimos", True, True),
        ("). Let every leader embrace the testing seasons of ministry with joy, allowing steadfast patience to finish its perfect work, so that we may stand approved before God as faithful workmen who need not to be ashamed.", False, False)
    ])
    sp()

    # Save docx files
    for path in docx_paths:
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        doc.save(path)
        print(f"[+] Successfully saved Word docx to: {path}")

    # Save md files
    full_md_text = "\n".join(md_lines)
    for path in md_paths:
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write(full_md_text)
        print(f"[+] Successfully saved Markdown to: {path}")

def count_words(doc_path):
    doc = Document(doc_path)
    total_words = 0
    for p in doc.paragraphs:
        words = p.text.split()
        total_words += len(words)
    return total_words

if __name__ == '__main__':
    base_dir = Path(r"c:\Users\Stephycopy\OneDrive\Desktop\Theos\theos")
    study_dir = base_dir / "Study group"

    docx_destinations = [
        base_dir / "Cell_Leaders_Conference_Faithfulness_in_Ministry_StudyGroup_Answers.docx",
        study_dir / "Cell_Leaders_Conference_Faithfulness_in_Ministry_StudyGroup_Answers.docx"
    ]

    md_destinations = [
        base_dir / "Cell_Leaders_Conference_Faithfulness_in_Ministry_StudyGroup_Answers.md",
        study_dir / "Cell_Leaders_Conference_Faithfulness_in_Ministry_StudyGroup_Answers.md"
    ]

    create_study_answers(docx_destinations, md_destinations)

    words = count_words(docx_destinations[0])
    print(f"[+] Verification: Total word count in Word document = {words} words.")
