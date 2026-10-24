#!/usr/bin/env python3
"""
format_faithfulness_in_ministry.py
Applies the Refined Claude Standard from the Transcript Formatting Guide:
- 100% preservation of speaker's authentic vocabulary and concepts
- Spoken syntax smoothed into clear, publication-grade expository prose
- Vocal fillers and oral clutter removed
- Pronoun drift harmonized into third person
- Scripture citations formatted into blockquotes with bold references
- Personal anecdotes and conversational remarks italicized
- Headings derived strictly from speaker's words
"""

from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def build_refined_document(docx_paths, md_paths):
    doc = Document()

    # Standard Letter (8.5" x 11.0"), 1.0" margins
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
    font.color.rgb = RGBColor(0x11, 0x18, 0x27)

    md_lines = []

    def title(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(18)
        md_lines.append(f"# {text}\n")

    def subtitle(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(12)
        run = p.add_run(text)
        run.bold = True
        run.italic = True
        run.font.name = 'Arial'
        run.font.size = Pt(14)
        md_lines.append(f"## *{text}*\n\n---\n")

    def h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"\n## {text}\n")
        return p

    def h3(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"\n### {text}\n")
        return p

    def script_ref(ref_str):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(ref_str)
        run.bold = True
        run.italic = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"> **{ref_str}**")
        return p

    def script_verse(verse_str):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        run = p.add_run(verse_str)
        run.bold = True
        run.italic = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"> {verse_str}")
        return p

    def sp():
        md_lines.append("")

    def p_body(runs):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        md_str = ""
        for text, bold, italic in runs:
            run = p.add_run(text)
            run.font.name = 'Arial'
            run.font.size = Pt(12)
            if bold:
                run.bold = True
            if italic:
                run.italic = True

            if bold and italic:
                md_str += f"***{text}***"
            elif bold:
                md_str += f"**{text}**"
            elif italic:
                md_str += f"*{text}*"
            else:
                md_str += text
        md_lines.append(md_str + "\n")
        return p

    def p_anecdote(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        run.italic = True
        md_lines.append(f"*{text}*\n")
        return p

    # =========================================================================
    # DOCUMENT TEXT — REFINED CLAUDE STANDARD
    # =========================================================================

    title("Cell Leaders Conference — Faithfulness in Ministry")
    subtitle("December 2021 — Teaching by Pastor Segun Ayodele")

    # --- SECTION 1 ---
    h2("It Is Required in Stewards That a Man Be Found Faithful")

    p_anecdote(
        "I shared this yesterday and I am going to run through it again. Many must have seen it on my Facebook; I rarely share my notes like that, but I wanted people to read and see it."
    )

    p_body([
        ("In 1 Corinthians 4, beginning from verse 1:", False, False)
    ])

    script_ref("1 Corinthians 4:1-2")
    script_verse("[1] Let a man so account of us, as of the ministers of Christ, and stewards of the mysteries of God.")
    script_verse("[2] Moreover it is required in stewards, that a man be found faithful.")
    sp()

    p_body([
        ("It is required that a man be found faithful. The word faithful is the same word for faith, ", False, False),
        ("pistos", True, True),
        (". It means to be reliable, to be trustworthy, and to be dependable.", False, False)
    ])

    p_body([
        ("Believers often say, 'God is faithful.' First Corinthians 10:13 affirms that God is faithful:", False, False)
    ])

    script_ref("1 Corinthians 10:13")
    script_verse("[13] There hath no temptation taken you but such as is common to man: but God is faithful, who will not suffer you to be tempted above that ye are able...")
    sp()

    p_body([
        ("God is faithful because He is reliable. He is always available. He does not alter His commitment, nor does He change His word in spite and despite opposing extremities. That is why faithfulness is also called loyalty, and loyalty means to be committed to a cause one hundred percent. Paul is saying that it is required of a steward to be found loyal—loyal to what one is doing, faithful, and dependable under every circumstance.", False, False)
    ])

    # --- SECTION 2 ---
    h2("Faithfulness vs. Ambition: Ruth and Naomi and the Test of Unfulfilled Desires")

    p_body([
        ("Consider Ruth and Naomi. When Ruth spoke to Naomi:", False, False)
    ])

    script_ref("Ruth 1:16")
    script_verse("[16] And Ruth said, Intreat me not to leave thee, or to return from following after thee: for whither thou goest, I will go; and where thou lodgest, I will lodge: thy people shall be my people, and thy God my God.")
    sp()

    p_body([
        ("Ruth was not looking for anything in return, and that is faithfulness.", False, False)
    ])

    p_body([
        ("There is a fundamental difference between faithfulness and ambition. One of the ways to know whether an individual is faithful or merely ambitious is when their personal desires are delayed or not given to them. There are people who join a church or enter a fellowship solely so that they can rise through the ranks. They put on an act of faithfulness, but at any slight alteration or denial of their personal desires, their attitude changes immediately. That sudden change proves they were never faithful.", False, False)
    ])

    p_body([
        ("Faithfulness has nothing to do with fluctuating circumstances. God gave His word, yet consider what He walked through before Christ died on the cross—it involved intense extremities and opposition. Faithful men do not break under pressure; they remain reliable, offering no excuses.", False, False)
    ])

    # --- SECTION 3 ---
    h2("Faithful Men Do Not Break Under Pressure: Elijah and Elisha")

    p_body([
        ("Consider Elijah and Elisha. Elisha did not know all that lay ahead; he simply followed Elijah faithfully. Historical records suggest Elisha came from a wealthy agricultural family and was likely older, yet he served Elijah with all his heart. When the king asked who possessed the word of the Lord, the answer given was:", False, False)
    ])

    script_ref("2 Kings 3:11")
    script_verse("[11] But Jehoshaphat said, Is there not here a prophet of the LORD, that we may enquire of the LORD by him? And one of the king of Israel's servants answered and said, Here is Elisha the son of Shaphat, which poured water on the hands of Elijah.")
    sp()

    p_body([
        ("Pouring water on the hands of Elijah was a term signifying that he served him faithfully. Elisha followed him through. As they pressed past Jericho toward the Jordan, the sons of the prophets stayed behind, asking, ", False, False),
        ("'Why are you following this man? What are you following him for? Do you not know he is going away today?'", False, True),
        (" Elisha did not answer their cynicism; he kept following. When Elijah asked, 'What do you ask of me?' Elisha replied, 'I want a double portion of your spirit,' which was the firstborn inheritance to head the company of the prophets.", False, False)
    ])

    p_body([
        ("Elijah told him, 'You have asked a hard thing; nevertheless, if you see me when I am taken from you, it shall be so.' And Elisha remained faithful.", False, False)
    ])

    p_anecdote(
        "One knows he was not acting for himself by what took place after Elijah was taken up and the mantle fell upon him. As he split the Jordan, he cried:"
    )

    script_ref("2 Kings 2:14")
    script_verse("[14] And he took the mantle of Elijah that fell from him, and smote the waters, and said, Where is the LORD God of Elijah? and when he also had smitten the waters, they parted hither and thither: and Elisha went over.")
    sp()

    p_body([
        ("Notice his cry: 'Where is the Lord God of Elijah?' He did not say, 'Where is the God of Elisha?' He made it known to everyone that he was wearing Elijah's ministry under the same God. That is a faithful person.", False, False)
    ])

    p_body([
        ("He was loyal, just like Joshua. Joshua stood upon the fact that he was executing the instructions of Moses, the servant of the Lord. Joshua never used Moses' human shortcomings to preach, saying, 'Moses could have brought us in, but he was not the one who eventually entered; you know what happened to him.' Joshua stayed strictly with Moses' instructions.", False, False)
    ])

    p_body([
        ("Loyal, faithful people do not act otherwise when leadership is absent. Faithful people conduct themselves in the service precisely as they would if Jesus were seated in the front row. Faithful people do in the service what they would do if their Pastor were seated right there. They do not act diligently merely because leadership is looking; in fact, they handle things behind leadership as though leadership were present. It is required in stewards that a man be found faithful, reliable, and trustworthy under every circumstance.", False, False)
    ])

    # --- SECTION 4 ---
    h2("Like a Broken Tooth and a Foot out of Joint")

    p_anecdote(
        "At the Believers Convention, Pastor Somi shared thoughts that highlighted Timothy's consistent faithfulness throughout Paul's writings."
    )

    p_body([
        ("One could readily call Paul's epistles Timothy's epistles, because Timothy was faithful and present everywhere.", False, False)
    ])

    p_body([
        ("In contrast, Proverbs 25:19 describes the pain of unfaithfulness:", False, False)
    ])

    script_ref("Proverbs 25:19")
    script_verse("[19] Confidence in an unfaithful man in time of trouble is like a broken tooth, and a foot out of joint.")
    sp()

    p_anecdote(
        "I have experienced this as well. A broken tooth is painful and cannot chew; a foot out of joint cannot walk."
    )

    p_body([
        ("When an individual has been given responsibility—whether as a cell leader, head of department, or pastor—their name may be on the roll, but if they are unfaithful, they are like a broken tooth: the church cannot chew and cannot walk, yet they occupy the position.", False, False)
    ])

    p_body([
        ("Anyone who claims that what one does in the service of God carries no consequences is deceiving people.", False, False)
    ])

    p_anecdote(
        "I told an individual last year: 'Read my lips. Before you were born, I was already teaching unconditional forgiveness. Before you were even eligible to go to hell, I was teaching the grace of God. But watch—read my lips: you will bear the consequences of your actions.' He had done nothing to me personally, but I told him: 'When you are serving God, you do not joke.'"
    )

    p_body([
        ("God practically treasures human lives. That is why the Lord Jesus declared:", False, False)
    ])

    script_ref("Luke 17:2")
    script_verse("[2] It were better for him that a millstone were hanged about his neck, and he cast into the sea, than that he should offend one of these little ones.")
    sp()

    # --- SECTION 5 ---
    h2("God Practically Treasures Lives: The Casual Handling of Disciples")

    p_anecdote(
        "A church member once came to me outlining plans to relocate abroad. I asked him, 'Have you thought about the people you disciple?' He replied casually, 'I was just thinking that since my pastor knows them, the church will take care of them.' You arrive in the United States, the United Kingdom, or Canada, and you start struggling in ministry, wondering why you are struggling—it is because back in Nigeria you treated God's people casually! I told him, 'No, you ought to have been thinking about these souls!'"
    )

    p_body([
        ("People often act as though discipleship is merely bringing people, dumping them into a service, and declaring that lives are being changed, until another personal opportunity changes them. A leader cannot think outside the people committed to their care. Whatever one does to human lives is how God evaluates their faithfulness.", False, False)
    ])

    p_body([
        ("Lives are like treasure; lives are like money. Most cell leaders have twenty or thirty people under their care; some have forty or fifty, and fellowship leaders have even more. Yet for those same twenty people, the leader claims they have no time, citing demanding jobs or distant residences, reeling out excuses enough for the Old Testament while praying for increase.", False, False)
    ])

    # --- SECTION 6 ---
    h2("Lives Are Like Money: The Akara Illustration")

    p_body([
        ("How can increase come under such negligence?", False, False)
    ])

    p_anecdote(
        "Consider the example I gave yesterday: Suppose you give someone one hundred thousand naira to start a business frying akara (bean cakes). The individual is ecstatic at receiving the capital. But two weeks, three weeks, and a month pass, and he begins sending text messages: 'Sorry, COVID came; people are wearing masks and not eating akara.' Three weeks later, he says, 'Beans is now expensive.' When asked for the one hundred thousand naira capital, he replies, 'The capital is gone, but to the glory of God, over the next year I can recover twenty thousand naira for you.' Then he attends a convention praying for one million naira! Anyone who gives him fifty thousand naira is acting foolishly, because look at his track record. Yet leaders do the same: ten souls were committed to their cell, nine vanished through neglect with only a cousin remaining, and they stand praying for increase. More of what?"
    )

    p_body([
        ("Lives are an investment. The person who preached the gospel to you invested, and your flock represents that continued investment. For every life a leader builds, other lives are impacted; but for every life a leader neglects, lives are being wasted.", False, False)
    ])

    p_body([
        ("Leaders must give daily reports on what they do concerning the lives under their charge. When a leader fails to report daily and offers excuses about depleted data or power outages, they are cultivating an unfaithful path.", False, False)
    ])

    p_anecdote(
        "If we were to announce an incentive of one thousand dollars per soul won, there would be no street left unreached—covering the earth would happen within a single month! Ten souls would mean ten thousand dollars, and immediately people would calculate the conversion rate. For one hundred souls, that would be one hundred thousand dollars—fifty million naira. People would ask, 'Pastor, why can't we hold outreaches every day?' No one would claim evening fatigue; they would declare, 'I am not tired in Jesus; I have the strength of God! What useless sleep am I sleeping? I am for the Lord!' Why? Because money is attached. Yet human lives are spiritual currency; they are infinitely expensive."
    )

    p_anecdote(
        "Recently, I suffered a very severe headache. Someone in the room suggested, 'Pastor, you should rest.' I replied, 'No. I know I should rest, but there is work to be done. I will look away from physical symptoms and act like the Word of God is true. The job has to be done.'"
    )

    # --- SECTION 7 ---
    h2("Faithful with Little: Success in Ministry Is Planned")

    p_body([
        ("As Jesus declared:", False, False)
    ])

    script_ref("Luke 16:10")
    script_verse("[10] He that is faithful in that which is least is faithful also in much: and he that is unjust in the least is unjust also in much.")
    sp()

    p_body([
        ("If a leader is not faithful with little, who will commit greater trust into their hands? People sometimes observe numerically large churches and assume their growth is merely the result of preaching false doctrine, using mantles, or distributing anointing oil. That is not the case. In many instances, those works expand because the people involved are totally dedicated, disciplined, and hardworking. Erroneous doctrine in itself does not build; dedication and discipline do.", False, False)
    ])

    p_anecdote(
        "My father always told me that success is planned. When he entered the university, knowing his reading attention span did not exceed one hour, he did not surrender to that limitation. Instead, he ensured that he read for one hour five separate times each day, resting between sessions. He cut short unnecessary social activities and graduated with first-class honors. He planned his success."
    )

    p_body([
        ("In the same way, ministry cannot grow merely because one is praying while neglecting practical execution. Praying without disciplined planning is like stockpiling ammunition in an armory without having a firearm to shoot it. Prayer provides the spiritual ammunition, but practical execution must discharge it.", False, False)
    ])

    p_anecdote(
        "I told our pastors that for every soul attending our conventions, while we may not account for visitors from other churches, we must give an accurate account of every single one of our own members. If we do not learn accountability with ten thousand, we will never manage one hundred thousand or one million. We must demonstrate accountability."
    )

    p_body([
        ("Faithfulness is the breath of ministry: it means always being at one's duty post.", False, False)
    ])

    # --- SECTION 8 ---
    h2("Always at the Duty Post: No Holidays from Ministry")

    p_body([
        ("The holiday season is not a holiday from ministry. A public holiday declared by the government does not suspend ministerial responsibility. If a businessman took a loan from the bank, interest runs through the holiday, even on Sunday morning while he is in church. How much more in spiritual ministry?", False, False)
    ])

    p_anecdote(
        "One Christmas Day, while returning from the United States through Dallas where Brother Pelumi was hosting me, we spent so much time greeting and talking that I missed my flight. Later, while traveling, I observed that security personnel were actively on duty, airplanes were operating, and air traffic controllers were directing flights without distraction. Doctors were on call delivering babies on Christmas Day, and police officers were inspecting vehicle trunks while greeting motorists with 'Merry Christmas.' Operations were running around the clock. I told myself: 'I am a preacher of the gospel; I must maintain spiritual alertness.'"
    )

    p_body([
        ("A minister must be faithful with time and faithful with alertness in all assigned duties.", False, False)
    ])

    # --- SECTION 9 ---
    h2("Testing and Proving: Let These Also First Be Proved")

    p_body([
        ("Circumstances will come to test an individual's faithfulness. When Samuel went to Bethlehem to anoint a king in 1 Samuel 16, he looked upon Eliab and assumed he was the Lord's anointed. But the Lord told Samuel that He had rejected Eliab. In the Hebrew text, the word carries the meaning of someone who failed the test.", False, False)
    ])

    p_body([
        ("One never knows when a test arrives. God tests people—not to destroy them, but the way an automobile is test-driven to verify its integrity. God boasted of Job: 'Have you considered My servant Job?' When trials come, James instructs:", False, False)
    ])

    script_ref("James 1:2-4")
    script_verse("[2] My brethren, count it all joy when ye fall into divers temptations;")
    script_verse("[3] Knowing this, that the trying of your faith worketh patience.")
    script_verse("[4] But let patience have her perfect work, that ye may be perfect and entire, wanting nothing.")
    sp()

    p_body([
        ("When an individual runs from tests, they run from the very character required for ministry.", False, False)
    ])

    p_body([
        ("In 1 Timothy 3, Paul writes concerning deacons—the associates of the pastor:", False, False)
    ])

    script_ref("1 Timothy 3:13")
    script_verse("[13] For they that have used the office of a deacon well purchase to themselves a good degree, and great boldness in the faith which is in Christ Jesus.")
    script_ref("1 Timothy 3:10")
    script_verse("[10] And let these also first be proved; then let them use the office of a deacon, being found blameless.")
    sp()

    p_body([
        ("To be proved means to be examined closely and tested thoroughly. Who does the proving? Paul does not say that the Holy Spirit will perform this proving in a mystical manner; rather, leadership must do it. In leadership, one must keep their eyes open to observe people and perceive their true character.", False, False)
    ])

    script_ref("2 Timothy 2:15")
    script_verse("[15] Study to shew thyself approved unto God, a workman that needeth not to be ashamed, rightly dividing the word of truth.")
    sp()

    p_body([
        ("The word 'approved' denotes one who has been tested. A leader cannot be untested; there must be proven experience through testing.", False, False)
    ])

    p_body([
        ("When the apostles in Acts 15 sent representatives to the Gentile churches, they commended Paul and Barnabas with these words:", False, False)
    ])

    script_ref("Acts 15:26")
    script_verse("[26] Men that have hazarded their lives for the name of our Lord Jesus Christ.")
    sp()

    p_body([
        ("They were chosen because they had been proven. They had put everything on the line, having stood firm through severe adversity rather than being fair-weather workers.", False, False)
    ])

    p_anecdote(
        "There are individuals who joined this church and quickly left because their objective was mere personal visibility. Even if someone arrived carrying the title of an Archbishop from elsewhere, they must pocket that title at the door, attend membership classes, and undergo testing in the Operations department, Works, or Security. Those who came desiring only to sing in the choir were assigned to Operations to clean the restrooms."
    )

    p_body([
        ("Workers must be proven. If someone came to serve rather than to seek a position, they will gladly serve wherever they are placed.", False, False)
    ])

    # --- SECTION 10 ---
    h2("Joseph vs. Gehazi: Faithfulness Across Every Circumstance")

    p_body([
        ("When people are tested, they emerge either as a Joseph or as a Gehazi.", False, False)
    ])

    p_body([
        ("Joseph was faithful as his father's favored son. When he was sold as a slave into Potiphar's house, no circumstance changed his integrity; he remained faithful, and Potiphar placed the entire estate under his hand. When Potiphar's wife solicited him, saying her husband would never know, Joseph declared that he would not sin against God. Even when she defamed him and had him cast into prison, he did not lapse into bitterness or plot a prison escape.", False, False)
    ])

    p_body([
        ("In prison, when the butler and the baker had dreams, Joseph was not resentful, complaining that dreams had caused his troubles. Instead, he faithfully interpreted their dreams. When the butler forgot him for two years, Joseph did not grumble. If an individual is going to manage a nation, they must know how to handle trouble.", False, False)
    ])

    p_body([
        ("When summoned before Pharaoh, Joseph did not set personal conditions demanding release, revenge against Potiphar's wife, or compensation. He simply declared that God would provide Pharaoh an answer of peace. Pharaoh recognized that the Spirit of God was in him, and placed him in charge of the nation.", False, False)
    ])

    p_anecdote(
        "Joseph was now in charge of the entire land, including Potiphar's household. When Potiphar's wife came seeking food provisions, Joseph did not retaliate; he provided rations and treated her with grace."
    )

    p_body([
        ("When his brothers arrived fearful after their father's death, manufacturing a message urging forgiveness, Joseph wept and comforted them:", False, False)
    ])

    script_ref("Genesis 50:20")
    script_verse("[20] But as for you, ye thought evil against me; but God meant it unto good, to bring to pass, as it is this day, to save much people alive.")
    sp()

    p_body([
        ("That reveals the character of Jesus Christ. No circumstance changes a faithful person.", False, False)
    ])

    p_body([
        ("In contrast, Gehazi allowed covetousness to rule him; he ran after Naaman for material gain and contracted leprosy. God does not use people without testing them. He tests workers to build stamina and commitment so that He can entrust more into their hands.", False, False)
    ])

    # --- SECTION 11 ---
    h2("Never Donate Responsibility to Unproven People")

    p_body([
        ("Leaders must never assign responsibility to people who have not been proven. Responsibility should never be handed out casually.", False, False)
    ])

    p_anecdote(
        "Someone once told me, 'Pastor, I am gathering some people in church.' I told him, 'Go and do that within your zone and submit your reports to the Zonal Pastor.' The person immediately became uncomfortable, and he stopped attending because his motive was simply to gain personal notice from the senior pastor."
    )

    p_body([
        ("Pastors must not alter this standard; workers must be observed directly.", False, False)
    ])

    p_anecdote(
        "If someone is appointed to a committee to supervise accommodations during a convention and the arrangement collapses, that is noted. If someone assigned to Works arrives late, that matters. If a leader had eight workers and six have disappeared with only two remaining, that matters significantly. Eye-service will not sustain anyone."
    )

    p_anecdote(
        "God permits certain fires so that those who cannot stand with leadership in the promised land do not cross the desert."
    )

    p_anecdote(
        "Furthermore, pastoral oversight cannot be exercised through social media. Following followers on an online page is not pastoring; there is no scriptural basis for pastoring through social media. Do not flee from testing; remain faithful."
    )

    # --- SECTION 12 ---
    h2("David the Shepherd: Keeping the Few Sheep")

    p_body([
        ("Consider David the shepherd. In Ezekiel 34, the prophetic name applied to Jesus is David. Look at the very first mention of David in 1 Samuel 16: when Samuel asked Jesse whether all his sons were present, Jesse replied:", False, False)
    ])

    script_ref("1 Samuel 16:11")
    script_verse("[11] And Samuel said unto Jesse, Are here all thy children? And he said, There remaineth yet the youngest, and, behold, he keepeth the sheep.")
    sp()

    p_body([
        ("The foundational testimony concerning David was that he was keeping the sheep.", False, False)
    ])

    p_body([
        ("When David went to the battlefront and his older brother Eliab mocked him concerning the few sheep left in the wilderness, David had already ensured that those sheep were left in the care of a keeper. When brought before King Saul, David cited his pastoral vigilance:", False, False)
    ])

    script_ref("1 Samuel 17:34-36")
    script_verse("[34] And David said unto Saul, Thy servant kept his father's sheep, and there came a lion, and a bear, and took a lamb out of the flock:")
    script_verse("[35] And I went out after him, and smote him, and delivered it out of his mouth: and when he arose against me, I caught him by his beard, and smote him, and slew him.")
    script_verse("[36] Thy servant slew both the lion and the bear...")
    sp()

    p_body([
        ("David did not lose a single lamb; he mirrored the Good Shepherd. He was faithful with the few. God proved David with the flock before entrusting him with the nation—just as Moses spent forty years tending sheep before leading Israel. Faithfulness must never be merely a thing of the past; faithfulness is what one is actively doing now.", False, False)
    ])

    # --- SECTION 13 ---
    h2("The Word and the Spirit: Resources to Multiply Lives")

    p_body([
        ("The Word of God and the Holy Spirit are the minister's resources to multiply what God has committed to their hands. The Word is what you preach, teach, and act upon. A leader is not called merely to give instructions in every direction, but to stand as an example:", False, False)
    ])

    script_ref("1 Timothy 4:12")
    script_verse("[12] Let no man despise thy youth; but be thou an example of the believers, in word, in conversation, in charity, in spirit, in faith, in purity.")
    script_ref("1 Peter 5:3")
    script_verse("[3] Neither as being lords over God's heritage, but being ensamples to the flock.")
    sp()

    p_body([
        ("A leader must listen to God's Word with 'the second hearing.' A leader does not listen merely for personal benefit; a leader listens on behalf of nations. That requires deep attention rather than casual listening. The Word of God and the Spirit are your resources; if you mishandle them, the work suffers. When teaching books are released, no one should need to compel a leader to acquire them.", False, False)
    ])

    p_anecdote(
        "Years ago, I would walk to Christian book distributors in Isolo, asking when new books by Kenneth Hagin or Frederick K.C. Price were arriving, booking copies down weeks in advance. When they began a national book club, I was among the top three in Nigeria as a teenager. Nobody had to compel me to purchase books."
    )

    p_anecdote(
        "When the Amplified Bible was extraordinarily rare, I would visit the only person I knew who owned a copy, sit down for thirty allotted minutes, and copy out the verses and cross-references by hand into my notebook. One must have that degree of spiritual hunger. Today, people can easily download recordings. There was a bookshop opposite Deeper Life where I purchased every cassette tape available by Brother Hagin and T.L. Osborn. In 1994, I went to a tape kiosk wearing slippers and requested every tape by Pastor Chris. The attendant looked at me and asked, 'Do you know what you are talking about? There are over two hundred tapes!' I replied that I wanted everything, and I purchased them all. That is dedication: you do not go halfway, and you do not simply display them on a shelf—you study them."
    )

    p_anecdote(
        "Twenty-four years ago, I began the Living Word Correspondence course, providing study books and examination questions to believers so they could study systematically. When you see our dedication to study today, it is not a recent development. Recently, a denomination asked me to review an 1,800-page curriculum. When they asked how long it would take, I requested four days. They asked, 'Did you say four months?' I said, 'Four days.' Within four days, I completed the review and returned it."
    )

    p_anecdote(
        "In 1994, God spoke to my heart: 'Do not follow people's results; follow their labor.' What are they doing? If they are fasting and praying, do it. If they are giving their all, do it. Ministry is not merely stepping onto a pulpit and telling people to raise their hands toward heaven."
    )

    p_anecdote(
        "In 1995 and 1996, when I lacked the money to buy books for everyone, I made photocopies of Brother Hagin's Bible study courses and distributed them during teaching meetings. What I am instructing you to do is not mere theory; give your all to acquire materials and invest in your spiritual life."
    )

    p_body([
        ("When listening to God's Word in conventions or church services, listen on behalf of nations. Take utterances and prayers personally. When the service ends, put what you heard into practice. Whenever God repeats something continuously, it indicates that it has not yet been practiced.", False, False)
    ])

    p_anecdote(
        "Many of you have demonstrated this consecration in the past: you used to walk long distances under the sun to attend meetings when you had no money for transit. Now, leaders excuse themselves because a vehicle broke down. There was a time in this ministry when confessing sickness was rejected; believers were compelled to stand upon divine healing."
    )

    p_anecdote(
        "We used to host Bible seminars on environmental sanitation Saturdays. Brethren would travel from Oyo State or Ogun State, arriving before 6:30 a.m. In the era of compact discs, believers would do anything to obtain a teaching disc. Today, recordings can be downloaded directly onto a phone; leaders must listen with headphones and feed upon the Word continuously."
    )

    p_anecdote(
        "Brother Toks is one of the most consistent men I know. If you wished to meet him, you simply stood before his door at five minutes before noon, and you would find him praying in tongues every single day. That is faithfulness, dedication, and consistency."
    )

    # --- SECTION 14 ---
    h2("The Four Demands of Pastoral Oversight")

    script_ref("Romans 15:20")
    script_verse("[20] Yea, so have I strived to preach the gospel, not where Christ was named, lest I should build upon another man's foundation.")
    sp()

    p_body([
        ("The practice of taking members from other assemblies must never be found among us. A disciple is someone who was unchurched, unsaved, and undiscipled—not someone dazzled with words or transferred from another fellowship. Leaders must raise believers from the foundation.", False, False)
    ])

    p_body([
        ("Four operational responsibilities are required of every cell leader:", False, False)
    ])

    h3("1. Soul Winning (Evangelism)")

    p_body([
        ("The flock must be won through active evangelism:", False, False)
    ])

    script_ref("Numbers 27:17")
    script_verse("[17] Which may go out before them, and which may go in before them, and which may lead them out, and which may bring them in; that the congregation of the LORD be not as sheep which have no shepherd.")
    script_ref("Ezekiel 34:11")
    script_verse("[11] For thus saith the Lord GOD; Behold, I, even I, will both search my sheep, and seek them out.")
    sp()

    p_body([
        ("If a meeting venue has empty seats, it reflects a diminishing passion for souls.", False, False)
    ])

    p_anecdote(
        "When Wizkid scheduled a concert in London at the O2 Arena, demand compelled organizers to extend it across three days. The people are out there in the city; we simply must labor to win them."
    )

    h3("2. Establish Them in the Local Church")

    p_body([
        ("Leaders must not merely invite people or dispatch them to church; a leader comes to church together with their disciples. Church attendance is an integral continuation of discipleship.", False, False)
    ])

    p_body([
        ("It is irresponsible for a leader to sit in the front rows while their disciples are scattered across the auditorium. Leaders must sit in the same area with their disciples to observe how they study, sing, worship, and take notes. Disciples should never sit through service without a Bible. Do not delegate that oversight to senior pastors; a cell leader is their pastor. That attentiveness is the very definition of pastoral oversight (", False, False),
        ("episkope", True, True),
        ("). Just as poor parenting produces disorder in society, careless pastoring produces undisciplined believers.", False, False)
    ])

    h3("3. Teach Them Daily")

    p_body([
        ("Leaders must have a structured plan ensuring that their flock feeds upon the Word of God daily. That does not mean physically meeting every day, but tracking their daily spiritual intake:", False, False)
    ])

    script_ref("Acts 20:31")
    script_verse("[31] Therefore watch, and remember, that by the space of three years I ceased not to warn every one night and day with tears.")
    script_ref("Acts 19:9-10")
    script_verse("[9] But when divers were hardened, and believed not, but spake evil of that way before the multitude, he departed from them, and separated the disciples, disputing daily in the school of one Tyrannus.")
    script_verse("[10] And this continued by the space of two years; so that all they which dwelt in Asia heard the word of the Lord Jesus, both Jews and Greeks.")
    sp()

    p_body([
        ("Paul taught daily in Ephesus. Cell leaders must actively follow up on what sermon series their flock is listening to, inquiring what was learned and ensuring continuous spiritual growth. Disciples are not merely invited to meetings; they are informed and brought together.", False, False)
    ])

    h3("4. Minister to Their Needs and Maintain Detailed Pastoral Records")

    script_ref("Acts 20:34-35")
    script_verse("[34] Yea, ye yourselves know, that these hands have ministered unto my necessities, and to them that were with me.")
    script_verse("[35] I have shewed you all things, how that so labouring ye ought to support the weak, and to remember the words of the Lord Jesus, how he said, It is more blessed to give than to receive.")
    sp()

    p_body([
        ("A leader must maintain accurate pastoral records of every individual under their care: full legal names, dates of birth, residential addresses, and occupations. If a leader cannot maintain diligent familiarity with twenty or fifty people, they cannot oversee hundreds or thousands.", False, False)
    ])

    # --- SECTION 15 ---
    h2("Men of Purpose Are Busy People")

    p_body([
        ("Pastoral ministry is a continuous, daily lifestyle. Even when a leader is not physically present, they are pastoring through prayer, planning, and communication. A leader must manage time intentionally, making room for family, vocational employment, personal communion, health, and flock oversight by removing frivolous activities.", False, False)
    ])

    p_anecdote(
        "Men of purpose are inherently busy people. I once visited a prominent Nigerian businessman in his office. He managed flights to Paris and New York with remarkable speed. On one occasion at 10:30 p.m., he invited me on his private jet to Ilorin to inspect an ongoing project; within twenty minutes of inspecting it, we flew back to Lagos before the driver waiting for me in the parking lot realized we had left."
    )

    p_body([
        ("People observe great results and assume they were achieved merely by passive confessions. Fulfillment in ministry is never achieved by mere confession; it demands relentless, disciplined labor.", False, False)
    ])

    p_anecdote(
        "A man of purpose is busy. Let us pray."
    )

    # --- EXPORT ---
    for docx_p in docx_paths:
        Path(docx_p).parent.mkdir(parents=True, exist_ok=True)
        doc.save(docx_p)
        print(f"[+] Saved Word docx to: {docx_p}")

    full_md = "\n".join(md_lines)
    for md_p in md_paths:
        Path(md_p).parent.mkdir(parents=True, exist_ok=True)
        with open(md_p, "w", encoding="utf-8") as f:
            f.write(full_md)
        print(f"[+] Saved Markdown to: {md_p}")

if __name__ == '__main__':
    base_dir = Path(r"c:\Users\Stephycopy\OneDrive\Desktop\Theos\theos")
    study_dir = base_dir / "Study group"

    docx_files = [
        base_dir / "Cell_Leaders_Conference_Faithfulness_in_Ministry.docx",
        study_dir / "Cell_Leaders_Conference_Faithfulness_in_Ministry.docx"
    ]

    md_files = [
        base_dir / "Cell_Leaders_Conference_Faithfulness_in_Ministry.md",
        study_dir / "Cell_Leaders_Conference_Faithfulness_in_Ministry.md"
    ]

    build_refined_document(docx_files, md_files)

    doc = Document(docx_files[0])
    total_words = sum(len(p.text.split()) for p in doc.paragraphs)
    print(f"[+] Verified: Total word count in Word document = {total_words} words.")
