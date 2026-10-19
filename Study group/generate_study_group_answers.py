#!/usr/bin/env python3
"""
generate_study_group_answers.py
Builds the comprehensive Study Group Answers for:
Cell Leaders Conference July 2021 – How Leadership changes you
STUDY GROUP QUESTIONS (5th – 11th Oct 2026)

Generates both .docx and .md versions following the exact compositional pattern,
intellectual rigor, and formatting standards of KeyIndicators_Track1_Answers.docx.
"""

import os
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def build_study_group_docx(output_path):
    doc = Document()

    # Page Margins: 1-inch all around, Letter size
    for sec in doc.sections:
        sec.page_width = Inches(8.5)
        sec.page_height = Inches(11.0)
        sec.top_margin = Inches(1.0)
        sec.bottom_margin = Inches(1.0)
        sec.left_margin = Inches(1.0)
        sec.right_margin = Inches(1.0)

    # Style defaults
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)
    font.color.rgb = RGBColor(0x11, 0x18, 0x27)

    def p_header(text, size=13, bold=True, italic=False, space_before=0, space_after=4):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(size)
        run.bold = bold
        run.italic = italic
        return p

    def p_qtitle(q_num):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(f"QUESTION {q_num}")
        run.font.name = 'Arial'
        run.font.size = Pt(13)
        run.bold = True
        return p

    def p_qtext(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(12)
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        run.bold = True
        return p

    def p_section_label(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        run.bold = True
        return p

    def p_subheading(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        run.bold = True
        return p

    def p_ref(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        run.bold = True
        run.italic = True
        return p

    def p_verse(text):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        run.bold = True
        run.italic = True
        return p

    def p_body(runs_data):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15
        for item in runs_data:
            if isinstance(item, tuple):
                r_text, r_bold, r_italic = item
            else:
                r_text, r_bold, r_italic = item, False, False
            run = p.add_run(r_text)
            run.font.name = 'Arial'
            run.font.size = Pt(12)
            run.bold = r_bold
            run.italic = r_italic
        return p

    # --- Header Information ---
    p_header("STUDY GROUP QUESTIONS (5th – 11th Oct 2026)", size=13, bold=True, space_before=0, space_after=2)
    p_header("Cell Leaders Conference July 2021 – How Leadership changes you", size=12, bold=True, italic=True, space_before=0, space_after=6)
    p_header("Name: AGBEDE STEPHEN AYOTOMIWA", size=12, bold=True, space_before=0, space_after=2)
    p_header("CHURCH: SAINTS COMMUNITY CHURCH, AKURE", size=12, bold=True, space_before=0, space_after=14)

    # =========================================================================
    # QUESTION 1
    # =========================================================================
    p_qtitle(1)
    p_qtext("Write a concise note on Leadership as seen in the scriptures and as seen in the world")

    p_section_label("INTRODUCTION")
    p_body([
        ("The local church is a place where a divine curriculum is followed, and central to that curriculum is the radical renewal of the believer's mind concerning spiritual realities. In no area is this renewal more urgently needed than in our conception of leadership. The human mind naturally operates by default within the categories of the world system. When believers and ministers hear the word 'leadership', they almost instinctively import the secular, political, and cultural definitions they have observed in society—definitions rooted in coercive authority, institutional hierarchy, and personal prominence. However, the scriptures do not merely adjust the world's concept of leadership; they completely dismantle and invert it. As Jesus explicitly declared in Matthew 20:25-26, ", False, False),
        ("Matthew 20:25-26", True, True),
        (", the model operating among the nations must never be the model operating in the kingdom of God: ", False, False),
        ("'it shall not be so among you.'", True, True),
        (" Biblical leadership is not an ascent to power over men; it is a descent into sacrificial service on behalf of men. To understand scriptural leadership, one must examine the fundamental contrast between the world's model of dominion and God's model of ministry, unpack the profound revelation of Jesus as the servant-son (", False, False),
        ("pais", True, True),
        ("), and understand how divine honor is commanded by service rather than demanded by office.", False, False)
    ])

    p_section_label("BODY")

    p_subheading("The World's Model of Leadership: Dominion, Coercion, and the Perks of Office")
    p_body([
        ("In the secular and political world, leadership is defined primarily by power, influence, and the apparatus of state authority. It is embodied in the police, the military, the judiciary, and governmental executive machinery. In this system, to be a leader is to possess the institutional capacity to enforce compliance, exercise jurisdiction, and project power over subordinates. Jesus pinpointed this exact mechanism in Matthew 20:25:", False, False)
    ])
    p_ref("Matthew 20:25")
    p_verse("[25] But Jesus called them unto him, and said, Ye know that the princes of the Gentiles exercise dominion over them, and they that are great exercise authority upon them.")
    p_body([
        ("The Gentile system operates on two principles: ", False, False),
        ("katakurieuo", True, True),
        (" (exercising lordship or dominion over people) and ", False, False),
        ("katexousiazo", True, True),
        (" (wielding heavy authority from above). In this secular paradigm, leadership is hierarchical, domineering, and self-serving. People seek leadership because of what they can extract from it: prestige, social elevation, public recognition, and the 'perks of office.'", False, False)
    ])
    p_body([
        ("The teaching exposes how this worldly mindset naturally infects human hearts—even among disciples who walked physically with Jesus. In Matthew 20:20-24, the mother of James and John came asking for the seats of honor on the right and left hand of Jesus in His kingdom. When the remaining ten disciples heard it, verse 24 records that ", False, False),
        ("'they were moved with indignation against the two brethren.'", True, True),
        (" The teaching makes a penetrating psychological observation: the ten were angry not because they possessed superior spiritual purity, but because they equally desired those seats of honor! Their indignation was the anger of competing ambition. They feared being eclipsed. In the worldly framework, leadership is a scarce commodity fought over for personal significance, where success is measured by how many people sit under you, how many serve you, and how high you sit in front.", False, False)
    ])

    p_subheading("The Scriptural Model of Leadership: Greatness through Servitude (Diakonia)")
    p_body([
        ("Jesus immediately arrested this worldly trajectory by establishing the foundational axiom of kingdom leadership:", False, False)
    ])
    p_ref("Matthew 20:26-28")
    p_verse("[26] But it shall not be so among you: but whosoever will be great among you, let him be your minister;")
    p_verse("[27] And whosoever will be chief among you, let him be your servant:")
    p_verse("[28] Even as the Son of man came not to be ministered unto, but to minister, and to give his life a ransom for many.")
    p_body([
        ("In the kingdom of God, greatness is measured not by how many people minister to you, but by whom you minister unto. The Greek word for 'minister' in verse 26 is ", False, False),
        ("diakonos", True, True),
        ("—a servant, an attendant, one who executes the commands of another and supplies the needs of others. The word for 'servant' in verse 27 is ", False, False),
        ("doulos", True, True),
        ("—a bondservant who has surrendered personal autonomy to serve the household. Jesus establishes that in God's idea, leadership is not a platform for self-aggrandizement; it is a posture of self-giving.", False, False)
    ])
    p_body([
        ("The teaching forcefully corrects the common religious misconception of a minister: some people imagine a minister is somebody who sits regally in front, receiving greetings, surrounded by protocol, and basking in privileges. That is completely worldly. A true minister is somebody who is prepared to give. What comes with the office is never personal entitlement; what comes with the office is what God has invested in the vessel solely so that the vessel can serve the people well. The office is an investment for stewardship, not an instrument for personal ambition.", False, False)
    ])

    p_subheading("Jesus as the Servant-Son (Pais): The Christological Benchmark in Acts 3, Isaiah 42, and Philippians 2")
    p_body([
        ("To provide the doctrinal anchor for this service, the teaching conducts an exegetical examination of the person of Jesus Christ. In Peter's sermon in Solomon's porch following the healing of the lame man, Peter proclaims:", False, False)
    ])
    p_ref("Acts 3:13")
    p_verse("[13] The God of Abraham, and of Isaac, and of Jacob, the God of our fathers, hath glorified his Son Jesus.")
    p_body([
        ("The teaching observes a vital lexical truth: the Greek word translated 'Son' in Acts 3:13 (and repeated in Acts 4:27 as 'holy child Jesus') is not the ordinary word ", False, False),
        ("huios", True, True),
        (" (a mature legal son) nor ", False, False),
        ("teknon", True, True),
        (" (a born child). It is the Greek word ", False, False),
        ("pais", True, True),
        (". The term ", False, False),
        ("pais", True, True),
        (" specifically denotes a servant who is a son—a son-servant. It describes a grown male servant who willingly yields his life and autonomy for divine service. This word directly links New Testament apostolic preaching to the messianic prophecy of Isaiah 42:1 in the Septuagint:", False, False)
    ])
    p_ref("Isaiah 42:1")
    p_verse("[1] Behold my servant, whom I uphold; mine elect, in whom my soul delighteth.")
    p_body([
        ("The Septuagint renders 'servant' here as ", False, False),
        ("pais", True, True),
        (". Jesus is the ultimate Servant-Son. He did not demand worship by brute divine prerogative; He proved His status by service. He gave Himself willingly, laying down His life as a ransom for many. Therefore, worship and honor belong to Him because of the depth of His sacrifice.", False, False)
    ])
    p_body([
        ("This theological reality is reinforced in Philippians 2:5-9:", False, False)
    ])
    p_ref("Philippians 2:5-7, 9")
    p_verse("[5] Let this mind be in you, which was also in Christ Jesus:")
    p_verse("[6] Who, being in the form of God, thought it not robbery to be equal with God:")
    p_verse("[7] But made himself of no reputation, and took upon him the form of a servant.")
    p_verse("[9] Wherefore God also hath highly exalted him, and given him a name which is above every name.")
    p_body([
        ("While believers frequently quote Philippians 2 to celebrate the authority of the believer or the exaltation of Christ's name, the teaching restores the passage to its precise literary and doctrinal context: Paul was commanding believers to possess the ", False, False),
        ("mind of Christ", True, True),
        ("—the mind of One who, although existing in the very form of God, emptied Himself and assumed the form of a bondservant (", False, False),
        ("doulos", True, True),
        ("). God highly exalted Him precisely because He lowered Himself to serve. The teaching delivers this profound theological verdict: ", False, False),
        ("service proves our divinity", True, True),
        (". The life of God within the believer is fundamentally a life of service. The life of the Spirit is a life of service. We are never more like God than when we are serving.", False, False)
    ])

    p_subheading("The Law of Honor: Commanded by Service, Never Demanded by Title")
    p_body([
        ("A critical corollary to this scriptural model is the proper understanding of honor in the local church. Without honor and respect for the vessel God is using, the people of God cannot receive what God has deposited in that vessel. Honor is explicitly instructed in Scripture. However, the teaching draws an unshakeable boundary: ", False, False),
        ("honor must never be demanded by title; honor is commanded by service.", True, True)
    ])
    p_body([
        ("The worldly leader demands deference based on their badge, their title, their political appointment, or their ecclesiastical rank. If people do not bow, the worldly leader takes offense and enforces subjugation. But the scriptural leader never demands honor. The scriptural leader allows their life, their speech, their generosity, their secret prayer life, their pastoral tears, and their daily sacrifice to command honor naturally. When a leader pours out their life to feed the flock, protect the sheep, visit the sick, and counsel the broken, honor gravitates to them spontaneously. Demanding honor by title is a sign of fleshly insecurity and worldly ambition; commanding honor by selfless service is the hallmark of Christ's minister.", False, False)
    ])

    p_section_label("CONCLUSION")
    p_body([
        ("In summary, leadership in the world and leadership in the scriptures represent two fundamentally incompatible kingdoms. The world's model is top-down, coercive, status-driven, and focused on extracting benefits from people through the exercise of dominion (", False, False),
        ("katakurieuo", True, True),
        ("). The scriptural model is Christ-centered, bottom-up, sacrificial, and focused on pouring out divine life for the enrichment of others through ministry (", False, False),
        ("diakonia", True, True),
        ("). Jesus, the supreme Son-Servant (", False, False),
        ("pais", True, True),
        ("), proved His divinity not by self-exaltation, but by making Himself of no reputation and dying on the cross. Therefore, the Christian leader—whether a pastor, cell leader, or departmental head—is not an ecclesiastical executive wielding authority over subordinates. He is a servant who has been entrusted with divine investments so that he may wash the feet of the saints. In the kingdom of God, the towel and the basin will forever remain the true symbols of spiritual authority.", False, False)
    ])

    # =========================================================================
    # QUESTION 2
    # =========================================================================
    p_qtitle(2)
    p_qtext("With the aid of this track, clearly explain to a fellow worker-in-training in church why he/she shouldn't quit or throw in the towel as touching serving or leading people via the work of the ministry, even if he/she has a moral failing or imperfections")

    p_section_label("INTRODUCTION")
    p_body([
        ("One of the most agonizing crises a worker-in-training or young leader faces in the local church occurs when their personal human weakness, character flaw, or moral failing collides with their ministerial responsibility. In that moment of failure, overwhelming condemnation, embarrassment, and guilt set in. The immediate instinct of the worker is to throw in the towel—to resign from their cell leadership, step down from the workforce, and retreat into isolation, concluding: 'I am a hypocrite. I am unworthy. I cannot serve God like this.'", False, False)
    ])
    p_body([
        ("However, this reaction is built on a fundamental misunderstanding of why God calls men into ministry in the first place. The teaching 'How Leadership Changes You' provides the definitive pastoral and theological antidote to this crisis. Through an exhaustive examination of Scripture, the teaching establishes that God does not enlist men because they are already perfect; God deliberately chooses raw, flawed, and imperfect human beings precisely so that the stream, discipline, and responsibility of leadership can transform them. Quitting in the face of failure is not humility; it is subtle pride running away from the very instrument God ordained for sanctification. To establish this truth for a struggling worker, we must examine the biblical pattern of the Twelve Apostles, unmask the pride behind quitting, analyze the coexistence of anointing and imperfection in the life of Samson, rest in the unchanging loyalty of God, and embrace the divine reality that under grace, a leader's mistake is redeemed to become their ministry.", False, False)
    ])

    p_section_label("BODY")

    p_subheading("The Biblical Blueprint: God Chooses Raw, Imperfect People to Transform Them")
    p_body([
        ("To steady the heart of the discouraged worker, one must first point them to the caliber of individuals Jesus chose to establish His church. The teaching emphasizes: ", False, False),
        ("'God always chooses normal human beings. Do not look for shining diamonds. God is going to use regular people.'", True, True),
        (" When you examine the Twelve Disciples in the Gospels, not a single one was chosen because they had achieved moral polish, doctrinal perfection, or emotional stability. Look at the biblical record:", False, False)
    ])
    p_body([
        ("1. ", True, False),
        ("Simon Peter: ", True, True),
        ("Andrew was the first person called by Jesus, and Andrew brought Peter (John 1:40-41). Andrew had persuasion and brought people into the church, but Jesus chose Peter to lead. Peter was crude, impulsive, boastful, and courageous to a fault. He constantly talked before he thought—a man of rush and impetuous speech who even dared to challenge and rebuke Jesus (Matthew 16:22). Even after the resurrection and the outpouring of the Spirit, Peter still exhibited traces of his old flaws: in Acts 5, Acts 8, Acts 15, and notably in Galatians 2:11-14, where Paul had to publicly confront Peter for ethnic hypocrisy and withdrawal from Gentile believers. Yet Jesus never stripped Peter of his apostleship. Over decades of bearing ministerial responsibility, leadership transformed him. By the time he wrote 1 and 2 Peter, he had matured into a seasoned, deeply humble patriarch exhorting elders not to lord it over God's heritage. God chose Peter the way he was—wrong—and used responsibility to transform him.", False, False)
    ])
    p_body([
        ("2. ", True, False),
        ("James and John: ", True, True),
        ("They were ambitious, political business partners nicknamed 'Sons of Thunder' because of their volatile tempers (wanting to call down fire on a Samaritan village in Luke 9:54). They schemed through their mother to secure chief seats in the kingdom. Yet Jesus chose them as young men and nurtured them until John became the Apostle of Love.", False, False)
    ])
    p_body([
        ("3. ", True, False),
        ("Nathanael: ", True, True),
        ("When Philip told him they had found the Messiah, Nathanael's first response was an unvarnished ethnic prejudice: ", False, False),
        ("'Can there any good thing come out of Nazareth?' (John 1:46).", True, True),
        (" He blurted out cultural bias without weighing his words. Yet Jesus called him an Israelite in whom was no guile.", False, False)
    ])
    p_body([
        ("4. ", True, False),
        ("Matthew the Publican: ", True, True),
        ("A tax collector, universally regarded by his Jewish society as a corrupt extortioner, a Roman collaborator, and a social outcast mixing with the wrong crowd. Yet Matthew had meticulous record-keeping skills, and Jesus redeemed that very capacity so that Matthew systematically organized the First Gospel.", False, False)
    ])
    p_body([
        ("5. ", True, False),
        ("Thomas, Simon the Zealot, and Judas: ", True, True),
        ("Thomas was an empirical skeptic who refused to believe without tangible physical evidence. Simon the Zealot was a militant political revolutionary with an aggressive, anti-Roman temper. Even Judas Iscariot was unfaithful with the purse, yet Jesus entrusted him with the treasury throughout His earthly ministry.", False, False)
    ])
    p_body([
        ("The worker-in-training must understand this vital principle: your imperfections and moral stumbles do not surprise God. He knew the entirety of your frailty before He appointed you. He did not call you because you were flawless; He called you so that the holy weight of serving others would force you to grow.", False, False)
    ])

    p_subheading("Unmasking the Temptation to Quit: Pride Masquerading as Guilt")
    p_body([
        ("The teaching delivers a devastating diagnosis of why workers throw in the towel after a moral failing or administrative collapse: ", False, False),
        ("dropping the ball is an act of PRIDE, not humility.", True, True)
    ])
    p_body([
        ("When a worker says, 'I fell into sin; I cannot hold this cell group; I cannot lead prayer; I am disqualifying myself,' they believe they are demonstrating deep spiritual remorse. But the teaching unmasks the reality: pride says, 'I cannot serve God unless I am perceived as flawless. I cannot endure the humiliation of having my moral vulnerability exposed. My self-image as a shining, upright leader has been tarnished, so I will take my toys and leave.' That is pride. You were never chosen because of your righteousness in the first place! You were chosen in spite of you.", False, False)
    ])
    p_body([
        ("True humility does not throw in the towel; true humility repents, receives the cleansing of Christ's blood, and says: ", False, False),
        ("'God gave me this responsibility so that I can get better. Now it is time to get back in the harness and get it right.'", True, True),
        (" When you abandon your post, you run away from the very sanctifying engine God ordained to cure your weakness. The teaching observes that nobody develops an intense prayer life in isolation. Your prayer life becomes intense because you carry people! Your fasting becomes serious because you have souls under your care! Your character is smoothed out because you must shepherd difficult sheep. When you quit, you step out of the river of the Spirit and revert to spiritual stagnation.", False, False)
    ])

    p_subheading("Samson as a Biblical Case Study: The Coexistence of Anointing and Imperfection")
    p_body([
        ("To illustrate that God functions through imperfect vessels while working on their character, the teaching examines the life of Samson:", False, False)
    ])
    p_ref("Judges 13:24-25")
    p_verse("[24] And the woman bare a son, and called his name Samson: and the child grew, and the LORD blessed him.")
    p_verse("[25] And the Spirit of the LORD began to move him at times in the camp of Dan between Zorah and Eshtaol.")
    p_body([
        ("Samson was set apart from his mother's womb as a Nazarite under specific consecration instructions. Yet his life was characterized by glaring moral missteps: in Judges 14:1-4, he demanded a Philistine wife from Timnath contrary to his parents' counsel (yet Scripture notes God sought an occasion against the Philistines); in Judges 15 and 16, he visited a harlot in Gaza. The Philistines surrounded the city to kill him at dawn, yet Samson arose at midnight, ripped up the massive city gates and posts, and carried them to the top of a hill!", False, False)
    ])
    p_body([
        ("Again and again, the supernatural anointing of God functioned in Samson's life despite his moral fragility. Why? Because the consecration and the covenant of God operated through him, and God was faithful to Israel. The anointing and human imperfection coexisted for a season because God is gracious. Then the teaching makes a monumental point: ", False, False),
        ("what eventually destroyed Samson's ministry was NOT his moral weakness—it was a lack of discretion and foolish disclosure.", True, True)
    ])
    p_body([
        ("Delilah pressed him daily until his soul was vexed unto death, and he revealed the secret of his heart (Judges 16:16-17). He weaponized relationships carelessly and threw away his consecration. The instruction to the struggling worker is profound: God's grace and anointing are patient with your weaknesses while He matures you, but you must cultivate discretion, guard your heart, protect your consecration, and refuse to allow the enemy to manipulate your vulnerabilities into total abandonment.", False, False)
    ])

    p_subheading("The Unchanging Loyalty of God: God Handles His Servants")
    p_body([
        ("The worker must be reminded of the sovereign loyalty of God. The teaching rings out with this comforting declaration: ", False, False),
        ("'No matter what men say about you, God is loyal to you. For the sake of His word, for the sake of His purpose and plan in the earth, He is loyal to you.'", True, True)
    ])
    p_body([
        ("Human beings are quick to write people off. Religious people love fault-finding. But the teaching warns believers against appointing themselves as judges over God's servants. In Numbers 12:1-10, when Moses married an Ethiopian woman, his siblings Miriam and Aaron murmured against him, saying, 'Has the Lord spoken only through Moses? Has He not spoken through us also?' They attacked his human choice and moral authority. But God immediately summoned them and rebuked them fiercely: ", False, False),
        ("'Wherefore then were ye not afraid to speak against my servant Moses?'", True, True),
        (" And Miriam was struck with leprosy! God knows how to handle His own servants. In Genesis 20, Abraham deceived Abimelech concerning Sarah, yet God warned Abimelech in a dream that Abraham was a prophet who would pray for him. God does not discard His servants when they stumble; He protects them, corrects them, and restores them.", False, False)
    ])
    p_body([
        ("Therefore, if men whisper about your failure or criticize your weakness, you must look away from the opinions of men and anchor your soul in the loyalty of God. He who called you is faithful. He will not cast you aside.", False, False)
    ])

    p_subheading("The Restoration Principle: The Mistake Becomes a Ministry")
    p_body([
        ("How does Jesus Himself respond when a key worker stumbles? In Luke 22:31-32, on the eve of Peter's catastrophic moral denial, Jesus addressed him:", False, False)
    ])
    p_ref("Luke 22:31-32")
    p_verse("[31] And the Lord said, Simon, Simon, behold, Satan hath desired to have you, that he may sift you as wheat:")
    p_verse("[32] But I have prayed for thee, that thy faith fail not: and when thou art converted, strengthen thy brethren.")
    p_body([
        ("Notice with awe the model of Christ: Jesus did not say, 'Simon, because you will deny me three times with curses in a few hours, you are hereby stripped of leadership and disqualified from the kingdom.' No! Jesus said: ", False, False),
        ("'I have prayed for you... and when you are turned back, STRENGTHEN YOUR BRETHREN.'", True, True),
        (" Jesus prepared Peter for recovery. He framed Peter's impending failure as a transition into deeper pastoral capacity. Under the redemptive mechanics of the cross, ", False, False),
        ("the mistake becomes a ministry!", True, True)
    ])
    p_body([
        ("The painful ground of failure where the worker stumbled will become the very territory of empathy, wisdom, and pastoral authority they use to rescue future disciples. A leader who has never fallen into weakness tends to be harsh, legalistic, and judgmental toward others. But a leader who knows what God accommodated in them will shepherd stumbling sheep with tender, restorative grace. As the African proverb quoted in the teaching states: ", False, False),
        ("'An elder who does not get hungry easily will have many children.'", True, True),
        (" A mature leader has a broad capacity to absorb human weakness and believe the best about raw people, because they know how much mercy they themselves received.", False, False)
    ])

    p_subheading("Embracing Process, Patience, and Separating Pain from the Assignment")
    p_body([
        ("Finally, the worker must be counselled concerning the necessity of process. The teaching identifies a dangerous generational flaw: the desire to 'chop sharp sharp'—the demand for quick results, instant perfection, and immediate spiritual elevation without enduring the furnace of process. Information without process does not produce wisdom. You can possess Bible knowledge and still behave foolishly, because process is what produces maturity.", False, False)
    ])
    p_body([
        ("Leadership is not a weekend workshop; it is a decades-long stream of the Spirit that reformats your character, your marriage, your speech, and your finances. The teacher noted: ", False, False),
        ("'I am not who I was over 30 years ago. I am not who I was 25 years ago. I am not who I was last year. And I am still leading. I am still changing.'", True, True),
        (" Growth takes time. The worker must be patient with themselves while staying accountable.", False, False)
    ])
    p_body([
        ("Crucially, the worker must learn to separate personal pain from the assignment: ", False, False),
        ("'Never bring your hurts to the pulpit. Clean your eyes before you get there. Never allow the wounds of one relationship to poison the next one you are building.'", True, True),
        (" You will carry burdens, yet you must lead. You are imperfect, yet you must minister to others. You are being perfected at the very same time that you are perfecting others! That paradox is the supernatural grace of God in operation.", False, False)
    ])

    p_section_label("CONCLUSION")
    p_body([
        ("To the worker-in-training standing at the verge of quitting, the verdict of the Word of God is clear and resounding: ", False, False),
        ("DO NOT THROW IN THE TOWEL.", True, True),
        (" God called you in spite of your flaws, and He is loyal to the covenant He made with you in Christ Jesus. Throwing in the towel is an escape into pride, while staying in ministry is an act of humble surrender to God's transforming grace. Your stumble is not your final chapter; it is the raw material from which God is forging your future message and pastoral tenderness.", False, False)
    ])
    p_body([
        ("Rise up today, shake off the dust of condemnation, wash your face, and re-engage the work of the ministry. Make this bold confession with your whole heart: ", False, False),
        ("'God is transforming me by leadership. My priorities are being reset. My cell leadership is changing me. My pastor is changing me. Those I am accountable to are changing me. I am being changed. I am changed. I will never stop being a leader. I will never get tired of doing it. I will never get bored of it. The Spirit of God is working on me in this season—and I am trusting Him more for my life, for my relationships, and for my ministry.'", True, True)
    ])

    doc.save(output_path)
    print(f"[+] Successfully generated DOCX: {output_path}")

def build_study_group_markdown(output_path):
    md_content = """# STUDY GROUP QUESTIONS (5th – 11th Oct 2026)
## *Cell Leaders Conference July 2021 – How Leadership changes you*

**Name:** AGBEDE STEPHEN AYOTOMIWA  
**CHURCH:** SAINTS COMMUNITY CHURCH, AKURE  

---

## QUESTION 1

**Write a concise note on Leadership as seen in the scriptures and as seen in the world**

### INTRODUCTION

The local church is a place where a divine curriculum is followed, and central to that curriculum is the radical renewal of the believer's mind concerning spiritual realities. In no area is this renewal more urgently needed than in our conception of leadership. The human mind naturally operates by default within the categories of the world system. When believers and ministers hear the word 'leadership', they almost instinctively import the secular, political, and cultural definitions they have observed in society—definitions rooted in coercive authority, institutional hierarchy, and personal prominence. However, the scriptures do not merely adjust the world's concept of leadership; they completely dismantle and invert it. As Jesus explicitly declared in Matthew 20:25-26, the model operating among the nations must never be the model operating in the kingdom of God: ***"it shall not be so among you."*** Biblical leadership is not an ascent to power over men; it is a descent into sacrificial service on behalf of men. To understand scriptural leadership, one must examine the fundamental contrast between the world's model of dominion and God's model of ministry, unpack the profound revelation of Jesus as the servant-son (***pais***), and understand how divine honor is commanded by service rather than demanded by office.

### BODY

#### The World's Model of Leadership: Dominion, Coercion, and the Perks of Office

In the secular and political world, leadership is defined primarily by power, influence, and the apparatus of state authority. It is embodied in the police, the military, the judiciary, and governmental executive machinery. In this system, to be a leader is to possess the institutional capacity to enforce compliance, exercise jurisdiction, and project power over subordinates. Jesus pinpointed this exact mechanism in Matthew 20:25:

***Matthew 20:25***
> ***[25]*** But Jesus called them unto him, and said, Ye know that the princes of the Gentiles exercise dominion over them, and they that are great exercise authority upon them.

The Gentile system operates on two principles: ***katakurieuo*** (exercising lordship or dominion over people) and ***katexousiazo*** (wielding heavy authority from above). In this secular paradigm, leadership is hierarchical, domineering, and self-serving. People seek leadership because of what they can extract from it: prestige, social elevation, public recognition, and the 'perks of office.'

The teaching exposes how this worldly mindset naturally infects human hearts—even among disciples who walked physically with Jesus. In Matthew 20:20-24, the mother of James and John came asking for the seats of honor on the right and left hand of Jesus in His kingdom. When the remaining ten disciples heard it, verse 24 records that ***"they were moved with indignation against the two brethren."*** The teaching makes a penetrating psychological observation: the ten were angry not because they possessed superior spiritual purity, but because they equally desired those seats of honor! Their indignation was the anger of competing ambition. They feared being eclipsed. In the worldly framework, leadership is a scarce commodity fought over for personal significance, where success is measured by how many people sit under you, how many serve you, and how high you sit in front.

#### The Scriptural Model of Leadership: Greatness through Servitude (*Diakonia*)

Jesus immediately arrested this worldly trajectory by establishing the foundational axiom of kingdom leadership:

***Matthew 20:26-28***
> ***[26]*** But it shall not be so among you: but whosoever will be great among you, let him be your minister;  
> ***[27]*** And whosoever will be chief among you, let him be your servant:  
> ***[28]*** Even as the Son of man came not to be ministered unto, but to minister, and to give his life a ransom for many.

In the kingdom of God, greatness is measured not by how many people minister to you, but by whom you minister unto. The Greek word for 'minister' in verse 26 is ***diakonos***—a servant, an attendant, one who executes the commands of another and supplies the needs of others. The word for 'servant' in verse 27 is ***doulos***—a bondservant who has surrendered personal autonomy to serve the household. Jesus establishes that in God's idea, leadership is not a platform for self-aggrandizement; it is a posture of self-giving.

The teaching forcefully corrects the common religious misconception of a minister: some people imagine a minister is somebody who sits regally in front, receiving greetings, surrounded by protocol, and basking in privileges. That is completely worldly. A true minister is somebody who is prepared to give. What comes with the office is never personal entitlement; what comes with the office is what God has invested in the vessel solely so that the vessel can serve the people well. The office is an investment for stewardship, not an instrument for personal ambition.

#### Jesus as the Servant-Son (*Pais*): The Christological Benchmark in Acts 3, Isaiah 42, and Philippians 2

To provide the doctrinal anchor for this service, the teaching conducts an exegetical examination of the person of Jesus Christ. In Peter's sermon in Solomon's porch following the healing of the lame man, Peter proclaims:

***Acts 3:13***
> ***[13]*** The God of Abraham, and of Isaac, and of Jacob, the God of our fathers, hath glorified his Son Jesus.

The teaching observes a vital lexical truth: the Greek word translated 'Son' in Acts 3:13 (and repeated in Acts 4:27 as 'holy child Jesus') is not the ordinary word ***huios*** (a mature legal son) nor ***teknon*** (a born child). It is the Greek word ***pais***. The term ***pais*** specifically denotes a servant who is a son—a son-servant. It describes a grown male servant who willingly yields his life and autonomy for divine service. This word directly links New Testament apostolic preaching to the messianic prophecy of Isaiah 42:1 in the Septuagint:

***Isaiah 42:1***
> ***[1]*** Behold my servant, whom I uphold; mine elect, in whom my soul delighteth.

The Septuagint renders 'servant' here as ***pais***. Jesus is the ultimate Servant-Son. He did not demand worship by brute divine prerogative; He proved His status by service. He gave Himself willingly, laying down His life as a ransom for many. Therefore, worship and honor belong to Him because of the depth of His sacrifice.

This theological reality is reinforced in Philippians 2:5-9:

***Philippians 2:5-7, 9***
> ***[5]*** Let this mind be in you, which was also in Christ Jesus:  
> ***[6]*** Who, being in the form of God, thought it not robbery to be equal with God:  
> ***[7]*** But made himself of no reputation, and took upon him the form of a servant.  
> ***[9]*** Wherefore God also hath highly exalted him, and given him a name which is above every name.

While believers frequently quote Philippians 2 to celebrate the authority of the believer or the exaltation of Christ's name, the teaching restores the passage to its precise literary and doctrinal context: Paul was commanding believers to possess the ***mind of Christ***—the mind of One who, although existing in the very form of God, emptied Himself and assumed the form of a bondservant (***doulos***). God highly exalted Him precisely because He lowered Himself to serve. The teaching delivers this profound theological verdict: ***service proves our divinity***. The life of God within the believer is fundamentally a life of service. The life of the Spirit is a life of service. We are never more like God than when we are serving.

#### The Law of Honor: Commanded by Service, Never Demanded by Title

A critical corollary to this scriptural model is the proper understanding of honor in the local church. Without honor and respect for the vessel God is using, the people of God cannot receive what God has deposited in that vessel. Honor is explicitly instructed in Scripture. However, the teaching draws an unshakeable boundary: ***honor must never be demanded by title; honor is commanded by service.***

The worldly leader demands deference based on their badge, their title, their political appointment, or their ecclesiastical rank. If people do not bow, the worldly leader takes offense and enforces subjugation. But the scriptural leader never demands honor. The scriptural leader allows their life, their speech, their generosity, their secret prayer life, their pastoral tears, and their daily sacrifice to command honor naturally. When a leader pours out their life to feed the flock, protect the sheep, visit the sick, and counsel the broken, honor gravitates to them spontaneously. Demanding honor by title is a sign of fleshly insecurity and worldly ambition; commanding honor by selfless service is the hallmark of Christ's minister.

### CONCLUSION

In summary, leadership in the world and leadership in the scriptures represent two fundamentally incompatible kingdoms. The world's model is top-down, coercive, status-driven, and focused on extracting benefits from people through the exercise of dominion (***katakurieuo***). The scriptural model is Christ-centered, bottom-up, sacrificial, and focused on pouring out divine life for the enrichment of others through ministry (***diakonia***). Jesus, the supreme Son-Servant (***pais***), proved His divinity not by self-exaltation, but by making Himself of no reputation and dying on the cross. Therefore, the Christian leader—whether a pastor, cell leader, or departmental head—is not an ecclesiastical executive wielding authority over subordinates. He is a servant who has been entrusted with divine investments so that he may wash the feet of the saints. In the kingdom of God, the towel and the basin will forever remain the true symbols of spiritual authority.

---

## QUESTION 2

**With the aid of this track, clearly explain to a fellow worker-in-training in church why he/she shouldn't quit or throw in the towel as touching serving or leading people via the work of the ministry, even if he/she has a moral failing or imperfections**

### INTRODUCTION

One of the most agonizing crises a worker-in-training or young leader faces in the local church occurs when their personal human weakness, character flaw, or moral failing collides with their ministerial responsibility. In that moment of failure, overwhelming condemnation, embarrassment, and guilt set in. The immediate instinct of the worker is to throw in the towel—to resign from their cell leadership, step down from the workforce, and retreat into isolation, concluding: 'I am a hypocrite. I am unworthy. I cannot serve God like this.'

However, this reaction is built on a fundamental misunderstanding of why God calls men into ministry in the first place. The teaching 'How Leadership Changes You' provides the definitive pastoral and theological antidote to this crisis. Through an exhaustive examination of Scripture, the teaching establishes that God does not enlist men because they are already perfect; God deliberately chooses raw, flawed, and imperfect human beings precisely so that the stream, discipline, and responsibility of leadership can transform them. Quitting in the face of failure is not humility; it is subtle pride running away from the very instrument God ordained for sanctification. To establish this truth for a struggling worker, we must examine the biblical pattern of the Twelve Apostles, unmask the pride behind quitting, analyze the coexistence of anointing and imperfection in the life of Samson, rest in the unchanging loyalty of God, and embrace the divine reality that under grace, a leader's mistake is redeemed to become their ministry.

### BODY

#### The Biblical Blueprint: God Chooses Raw, Imperfect People to Transform Them

To steady the heart of the discouraged worker, one must first point them to the caliber of individuals Jesus chose to establish His church. The teaching emphasizes: ***"God always chooses normal human beings. Do not look for shining diamonds. God is going to use regular people."*** When you examine the Twelve Disciples in the Gospels, not a single one was chosen because they had achieved moral polish, doctrinal perfection, or emotional stability. Look at the biblical record:

1. **Simon Peter:** Andrew was the first person called by Jesus, and Andrew brought Peter (John 1:40-41). Andrew had persuasion and brought people into the church, but Jesus chose Peter to lead. Peter was crude, impulsive, boastful, and courageous to a fault. He constantly talked before he thought—a man of rush and impetuous speech who even dared to challenge and rebuke Jesus (Matthew 16:22). Even after the resurrection and the outpouring of the Spirit, Peter still exhibited traces of his old flaws: in Acts 5, Acts 8, Acts 15, and notably in Galatians 2:11-14, where Paul had to publicly confront Peter for ethnic hypocrisy and withdrawal from Gentile believers. Yet Jesus never stripped Peter of his apostleship. Over decades of bearing ministerial responsibility, leadership transformed him. By the time he wrote 1 and 2 Peter, he had matured into a seasoned, deeply humble patriarch exhorting elders not to lord it over God's heritage. God chose Peter the way he was—wrong—and used responsibility to transform him.

2. **James and John:** They were ambitious, political business partners nicknamed 'Sons of Thunder' because of their volatile tempers (wanting to call down fire on a Samaritan village in Luke 9:54). They schemed through their mother to secure chief seats in the kingdom. Yet Jesus chose them as young men and nurtured them until John became the Apostle of Love.

3. **Nathanael:** When Philip told him they had found the Messiah, Nathanael's first response was an unvarnished ethnic prejudice: ***"Can there any good thing come out of Nazareth?" (John 1:46).*** He blurted out cultural bias without weighing his words. Yet Jesus called him an Israelite in whom was no guile.

4. **Matthew the Publican:** A tax collector, universally regarded by his Jewish society as a corrupt extortioner, a Roman collaborator, and a social outcast mixing with the wrong crowd. Yet Matthew had meticulous record-keeping skills, and Jesus redeemed that very capacity so that Matthew systematically organized the First Gospel.

5. **Thomas, Simon the Zealot, and Judas:** Thomas was an empirical skeptic who refused to believe without tangible physical evidence. Simon the Zealot was a militant political revolutionary with an aggressive, anti-Roman temper. Even Judas Iscariot was unfaithful with the purse, yet Jesus entrusted him with the treasury throughout His earthly ministry.

The worker-in-training must understand this vital principle: your imperfections and moral stumbles do not surprise God. He knew the entirety of your frailty before He appointed you. He did not call you because you were flawless; He called you so that the holy weight of serving others would force you to grow.

#### Unmasking the Temptation to Quit: Pride Masquerading as Guilt

The teaching delivers a devastating diagnosis of why workers throw in the towel after a moral failing or administrative collapse: ***dropping the ball is an act of PRIDE, not humility.***

When a worker says, 'I fell into sin; I cannot hold this cell group; I cannot lead prayer; I am disqualifying myself,' they believe they are demonstrating deep spiritual remorse. But the teaching unmasks the reality: pride says, 'I cannot serve God unless I am perceived as flawless. I cannot endure the humiliation of having my moral vulnerability exposed. My self-image as a shining, upright leader has been tarnished, so I will take my toys and leave.' That is pride. You were never chosen because of your righteousness in the first place! You were chosen in spite of you.

True humility does not throw in the towel; true humility repents, receives the cleansing of Christ's blood, and says: ***"God gave me this task so that I can get better. Now it is time to get back in the harness and get it right."*** When you abandon your post, you run away from the very sanctifying engine God ordained to cure your weakness. The teaching observes that nobody develops an intense prayer life in isolation. Your prayer life becomes intense because you carry people! Your fasting becomes serious because you have souls under your care! Your character is smoothed out because you must shepherd difficult sheep. When you quit, you step out of the river of the Spirit and revert to spiritual stagnation.

#### Samson as a Biblical Case Study: The Coexistence of Anointing and Imperfection

To illustrate that God functions through imperfect vessels while working on their character, the teaching examines the life of Samson:

***Judges 13:24-25***
> ***[24]*** And the woman bare a son, and called his name Samson: and the child grew, and the LORD blessed him.  
> ***[25]*** And the Spirit of the LORD began to move him at times in the camp of Dan between Zorah and Eshtaol.

Samson was set apart from his mother's womb as a Nazarite under specific consecration instructions. Yet his life was characterized by glaring moral missteps: in Judges 14:1-4, he demanded a Philistine wife from Timnath contrary to his parents' counsel (yet Scripture notes God sought an occasion against the Philistines); in Judges 15 and 16, he visited a harlot in Gaza. The Philistines surrounded the city to kill him at dawn, yet Samson arose at midnight, ripped up the massive city gates and posts, and carried them to the top of a hill!

Again and again, the supernatural anointing of God functioned in Samson's life despite his moral fragility. Why? Because the consecration and the covenant of God operated through him, and God was faithful to Israel. The anointing and human imperfection coexisted for a season because God is gracious. Then the teaching makes a monumental point: ***what eventually destroyed Samson's ministry was NOT his moral weakness—it was a lack of discretion and foolish disclosure.***

Delilah pressed him daily until his soul was vexed unto death, and he revealed the secret of his heart (Judges 16:16-17). He weaponized relationships carelessly and threw away his consecration. The instruction to the struggling worker is profound: God's grace and anointing are patient with your weaknesses while He matures you, but you must cultivate discretion, guard your heart, protect your consecration, and refuse to allow the enemy to manipulate your vulnerabilities into total abandonment.

#### The Unchanging Loyalty of God: God Handles His Servants

The worker must be reminded of the sovereign loyalty of God. The teaching rings out with this comforting declaration: ***"No matter what men say about you, God is loyal to you. For the sake of His word, for the sake of His purpose and plan in the earth, He is loyal to you."***

Human beings are quick to write people off. Religious people love fault-finding. But the teaching warns believers against appointing themselves as judges over God's servants. In Numbers 12:1-10, when Moses married an Ethiopian woman, his siblings Miriam and Aaron murmured against him, saying, 'Has the Lord spoken only through Moses? Has He not spoken through us also?' They attacked his human choice and moral authority. But God immediately summoned them and rebuked them fiercely: ***"Wherefore then were ye not afraid to speak against my servant Moses?"*** And Miriam was struck with leprosy! God knows how to handle His own servants. In Genesis 20, Abraham deceived Abimelech concerning Sarah, yet God warned Abimelech in a dream that Abraham was a prophet who would pray for him. God does not discard His servants when they stumble; He protects them, corrects them, and restores them.

Therefore, if men whisper about your failure or criticize your weakness, you must look away from the opinions of men and anchor your soul in the loyalty of God. He who called you is faithful. He will not cast you aside.

#### The Restoration Principle: The Mistake Becomes a Ministry

How does Jesus Himself respond when a key worker stumbles? In Luke 22:31-32, on the eve of Peter's catastrophic moral denial, Jesus addressed him:

***Luke 22:31-32***
> ***[31]*** And the Lord said, Simon, Simon, behold, Satan hath desired to have you, that he may sift you as wheat:  
> ***[32]*** But I have prayed for thee, that thy faith fail not: and when thou art converted, strengthen thy brethren.

Notice with awe the model of Christ: Jesus did not say, 'Simon, because you will deny me three times with curses in a few hours, you are hereby stripped of leadership and disqualified from the kingdom.' No! Jesus said: ***"I have prayed for you... and when you are turned back, STRENGTHEN YOUR BRETHREN."*** Jesus prepared Peter for recovery. He framed Peter's impending failure as a transition into deeper pastoral capacity. Under the redemptive mechanics of the cross, ***the mistake becomes a ministry!***

The painful ground of failure where the worker stumbled will become the very territory of empathy, wisdom, and pastoral authority they use to rescue future disciples. A leader who has never fallen into weakness tends to be harsh, legalistic, and judgmental toward others. But a leader who knows what God accommodated in them will shepherd stumbling sheep with tender, restorative grace. As the African proverb quoted in the teaching states: ***"An elder who does not get hungry easily will have many children."*** A mature leader has a broad capacity to absorb human weakness and believe the best about raw people, because they know how much mercy they themselves received.

#### Embracing Process, Patience, and Separating Pain from the Assignment

Finally, the worker must be counselled concerning the necessity of process. The teaching identifies a dangerous generational flaw: the desire to 'chop sharp sharp'—the demand for quick results, instant perfection, and immediate spiritual elevation without enduring the furnace of process. Information without process does not produce wisdom. You can possess Bible knowledge and still behave foolishly, because process is what produces maturity.

Leadership is not a weekend workshop; it is a decades-long stream of the Spirit that reformats your character, your marriage, your speech, and your finances. The teacher noted: ***"I am not who I was over 30 years ago. I am not who I was 25 years ago. I am not who I was last year. And I am still leading. I am still changing."*** Growth takes time. The worker must be patient with themselves while staying accountable.

Crucially, the worker must learn to separate personal pain from the assignment: ***"Never bring your hurts to the pulpit. Clean your eyes before you get there. Never allow the wounds of one relationship to poison the next one you are building."*** You will carry burdens, yet you must lead. You are imperfect, yet you must minister to others. You are being perfected at the very same time that you are perfecting others! That paradox is the supernatural grace of God in operation.

### CONCLUSION

To the worker-in-training standing at the verge of quitting, the verdict of the Word of God is clear and resounding: ***DO NOT THROW IN THE TOWEL.*** God called you in spite of your flaws, and He is loyal to the covenant He made with you in Christ Jesus. Throwing in the towel is an escape into pride, while staying in ministry is an act of humble surrender to God's transforming grace. Your stumble is not your final chapter; it is the raw material from which God is forging your future message and pastoral tenderness.

Rise up today, shake off the dust of condemnation, wash your face, and re-engage the work of the ministry. Make this bold confession with your whole heart:

> ***"God is transforming me by leadership. My priorities are being reset. My cell leadership is changing me. My pastor is changing me. Those I am accountable to are changing me. I am being changed. I am changed. I will never stop being a leader. I will never get tired of doing it. I will never get bored of it. The Spirit of God is working on me in this season—and I am trusting Him more for my life, for my relationships, and for my ministry."***
"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(md_content)
    print(f"[+] Successfully generated Markdown: {output_path}")

if __name__ == '__main__':
    base_dir = Path(r"c:\Users\Stephycopy\OneDrive\Desktop\Theos\theos")
    study_group_dir = base_dir / "Study group"
    study_group_dir.mkdir(parents=True, exist_ok=True)

    # 1. Output files in 'Study group' folder
    sg_docx = study_group_dir / "Cell_Leaders_Conference_How_Leadership_Changes_You_StudyGroup_Answers.docx"
    sg_md = study_group_dir / "Cell_Leaders_Conference_How_Leadership_Changes_You_StudyGroup_Answers.md"

    # 2. Output files in root of theos for quick access
    root_docx = base_dir / "Cell_Leaders_Conference_How_Leadership_Changes_You_StudyGroup_Answers.docx"
    root_md = base_dir / "Cell_Leaders_Conference_How_Leadership_Changes_You_StudyGroup_Answers.md"

    build_study_group_docx(sg_docx)
    build_study_group_markdown(sg_md)

    build_study_group_docx(root_docx)
    build_study_group_markdown(root_md)

    # Count words
    doc = Document(sg_docx)
    words = sum(len(p.text.split()) for p in doc.paragraphs)
    print(f"[+] Word count in generated Study Group Answers: {words} words.")
