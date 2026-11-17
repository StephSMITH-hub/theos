# -*- coding: utf-8 -*-
"""
Master Builder for:
In the Spirit by the Spirit – Power Gifts - Track 2
Structured Verbatim Transcript with Provided Bible Verses (KJV)

Matches the exact repository standard of:
- 'In the spirit by the spirit 7b Laying on of hands Track 2.docx'
- 'In the spirit by the spirit 7b Laying on of hands Track 5.docx'

Key Features:
1. Preserves 100% of spoken words without altering wording (oral flow, banter, Greek words, jokes, interactions, prayer).
2. Spoken discourse is broken down into clean, short, punchy paragraphs (1-2 sentences each).
3. Every cited/read scripture provides the reference heading in BOLD and full KJV verses with bracketed numbers '[#]' in BOLD on separate lines.
4. Styled with Lexend/Calibri, 1.0 inch margins, clear spacing after paragraphs.
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor

# Data elements: tuple of (element_type, text)
# element_type: 'p' (spoken discourse), 'ref' (scripture reference heading), 'v' (scripture verse line)

ELEMENTS = [
    # Opening
    ('p', "Open your word this morning, that our minds receive wisdom, that everyone be wiser upon hearing the truth of your word this morning, in the name of Jesus. Amen."),
    ('p', "Oh, praise the Lord!"),
    ('p', "2 Corinthians 1."),
    ('p', "In the Spirit, by the Spirit, this series started last year March, I think, and we've been looking at the things of the Spirit generally."),
    ('p', "Sometimes, spiritual things could take a long time, longer than you thought, but we're blessed anyway."),
    ('p', "Verse 18, 19 rather:"),
    ('ref', "2 Corinthians 1:18-22"),
    ('v', "[18] But as God is true, our word toward you was not yea and nay."),
    ('v', "[19] For the Son of God, Jesus Christ, who was preached among you by us, even by me and Silvanus and Timotheus, was not yea and nay, but in him was yea."),
    ('v', "[20] For all the promises of God in him are yea, and in him Amen, unto the glory of God by us."),
    ('v', "[21] Now he which stablisheth us with you in Christ, and hath anointed us, is God;"),
    ('v', "[22] Who hath also sealed us, and given the earnest of the Spirit in our hearts."),
    ('p', "Can you say amen?"),
    ('p', "Romans 15:8, he says:"),
    ('ref', "Romans 15:8"),
    ('v', "[8] Now I say that Jesus Christ was a minister of the circumcision for the truth of God, to confirm the promises made unto the fathers:"),
    ('p', "Praise the Lord!"),
    ('p', "And so we confirmed as what established."),
    ('p', "And so we said, quickly let's just move faster, that salvation, when you say salvation and you say the Spirit, right, you're saying the same thing."),
    ('p', "There cannot be salvation without the giving of the Spirit. The giving of the Spirit is salvation. Are you listening?"),
    ('p', "Now, there is a study we're probably going to do in the second service that may help our understanding, because you will find the word 'giving of the Spirit' used in the Book of Acts that seems to differ from what I just said."),
    ('p', "But anyway, salvation is synonymous with the giving of the Spirit."),
    ('p', "That is, you cannot be saved and not have the Spirit of God. Salvation is the giving, we can say the gift of the Spirit is salvation."),

    # Salvation & Resurrection
    ('p', "Because now, I want you to pay attention to something, and you pay attention to everything. I said they are synonymous terms."),
    ('p', "Now you know last week we looked about Jesus, we said Jesus is sacrifice, and we said most of the time our minds always stop at the cross. Remember I said that?"),
    ('p', "You know I said that, that we think Jesus died for my sins, and really the death of Jesus is one of the events of redemption."),
    ('p', "I'll take that again: the death of Jesus is one of the events of redemption."),
    ('p', "His death, now listen carefully now, his death is sacrificial, but his death is not the sacrifice for salvation."),
    ('p', "I'll take that again: his death is sacrificial, in other words he suffered, but his death is not the sacrifice for salvation."),
    ('p', "Make sure that you listen to me well for this, so that you don't think I said what I didn't say, right?"),
    ('p', "Now, we said, and I said last week, that what Jesus did is evidenced by his Spirit that he gave us."),
    ('p', "Now let me go back to what I said earlier on, that his death is one of the events."),
    ('p', "There are three events that mark the gospel of Christ, three: 1 Corinthians 15:3 and 4 states them. He said:"),
    ('ref', "1 Corinthians 15:3-4"),
    ('v', "[3] For I delivered unto you first of all that which I also received, how that Christ died for our sins according to the scriptures;"),
    ('v', "[4] And that he was buried, and that he rose again the third day according to the scriptures:"),
    ('p', "Notice that he seems to, that Paul seems to, okay he died, okay, like one event on its own, then says he was buried and raised again the third day according to the scriptures."),
    ('p', "He didn't say he was buried according to the scriptures, then he was raised again according to the scriptures. Are you listening? Are you listening?"),
    ('p', "Because if you die, really you'll be buried anyway. That he was buried was proof that he died, okay?"),
    ('p', "That he rose from the dead, okay, means that he was buried, you get that? You understand that?"),
    ('p', "Okay, so he has two or three events, right, that mark the gospel."),
    ('p', "So the gospel is not that Christ died; the gospel is Christ died, he was buried, and he rose from the dead."),
    ('p', "So are you listening? Are you listening? And that's how Paul puts it."),

    # 1 Cor 15
    ('p', "Now, notice something in 1 Corinthians 15, go there quickly. 1 Corinthians 15 verse 14."),
    ('p', "Now he says, some folks are arguing whether Jesus rose from the dead or not, or sorry, how the dead rise, sorry, let me start from there."),
    ('p', "And he says:"),
    ('ref', "1 Corinthians 15:14"),
    ('v', "[14] And if Christ be not risen, then is our preaching vain, and your faith is also vain."),
    ('p', "And he says, 'And if Christ be not what?' 14, 'be not what? Risen, okay, let's go on, our preaching is vain, and your faith is also vain.'"),
    ('p', "That word vain is fruitless, temporary."),
    ('p', "Now look at 17, let's take 17 together. 'And if Christ,' everybody let's read it together, let's go:"),
    ('ref', "1 Corinthians 15:17"),
    ('v', "[17] And if Christ be not raised, your faith is vain; ye are yet in your sins."),
    ('p', "Now he said, if Christ wasn't raised from the dead, your faith is vain, and you are still a sinner. He didn't say if Christ did not die. Are you listening? Are you listening?"),
    ('p', "So he says wasn't raised from the dead. So the rise or the raising or the resurrection of Jesus from the dead is the crown of redemption. It's the proof of redemption. It's the proof that man had been redeemed."),
    ('p', "So you need to listen to that very well. If Christ wasn't raised from the dead, then you are yet in your sins."),
    ('p', "Praise the name of Jesus. Hallelujah."),

    # John 3
    ('p', "Okay, now let's quickly, we're going to move like this to come back this way, you must be ready to move that way."),
    ('p', "Now in John 3, look at the words of Jesus. In John 3:3, verse 6, I knew that verse way before I got saved. Verily, verily, I say unto you, I used to wonder what verily verily meant, verily verily. Verily, verily, I say unto you, you there?"),
    ('ref', "John 3:3"),
    ('v', "[3] Jesus answered and said unto him, Verily, verily, I say unto thee, Except a man be born again, he cannot see the kingdom of God."),
    ('p', "Now notice, except a man be born again, he cannot see the kingdom of God."),
    ('p', "So Nicodemus asks in verse 4, what does it mean to be born again, that's basically what Nicodemus was saying. And look at what Jesus said:"),
    ('ref', "John 3:5-8"),
    ('v', "[5] Jesus answered, Verily, verily, I say unto thee, Except a man be born of water and of the Spirit, he cannot enter into the kingdom of God."),
    ('v', "[6] That which is born of the flesh is flesh; and that which is born of the Spirit is spirit."),
    ('v', "[7] Marvel not that I said unto thee, Ye must be born again."),
    ('v', "[8] The wind bloweth where it listeth, and thou hearest the sound thereof, but canst not tell whence it cometh, and whither it goeth: so is every one that is born of the Spirit."),
    ('p', "So notice he uses the word 'born of the Spirit' how many times? Three times, right? So come on, how many times?"),
    ('p', "So born of the Spirit explains born again. Is that clear? Is that clear? But, come on, is that clear? Right."),
    ('p', "So that means, you know, he says you have to be born again. Now you don't give birth to yourself, and you don't determine that anyway, it's the parent that determines that anyway."),
    ('p', "So he says be born of the Spirit. So Jesus calls it born of the Spirit."),

    # John 4 & John 7
    ('p', "In John 4:14, he puts it differently. He says:"),
    ('ref', "John 4:14"),
    ('v', "[14] But whosoever drinketh of the water that I shall give him shall never thirst; but the water that I shall give him shall be in him a well of water springing up into everlasting life."),
    ('p', "Then again John 7:38, 39."),
    ('p', "So Jesus in John 3 says born of the Spirit. Now in John 4:14, he says I'll give him, right, I'll give him, that is, I'll give to whoever believes, that's the word whoever drinks."),
    ('p', "I'll give to him water, okay, that springs up unto everlasting life. I'll give him, okay, so if a man believes I'll give him, that is, he will have water like a well springing up unto everlasting life. So the first one says he's born, the second one says I'll give him."),
    ('p', "Now John 7, John 7 in verse 37:"),
    ('ref', "John 7:37-39"),
    ('v', "[37] In the last day, that great day of the feast, Jesus stood and cried, saying, If any man thirst, let him come unto me, and drink."),
    ('v', "[38] He that believeth on me, as the scripture hath said, out of his belly shall flow rivers of living water."),
    ('v', "[39] (But this spake he of the Spirit, which they that believe on him should receive: for the Holy Ghost was not yet given; because that Jesus was not yet glorified.)"),
    ('p', "So that means in salvation, listen, there is a new birth, right, and then something is received. You don't just get saved the way you are."),
    ('p', "That is, there is what is given. God gives something. And Jesus says he gives the Spirit, which is the same thing as you're born of the Spirit."),

    # John 14 - Indwelling Spirit
    ('p', "Now go to John 14. John 14 and verse 16. He says:"),
    ('ref', "John 14:16-17"),
    ('v', "[16] And I will pray the Father, and he shall give you another Comforter, that he may abide with you for ever;"),
    ('v', "[17] Even the Spirit of truth; whom the world cannot receive, because it seeth him not, neither knoweth him: but ye know him; for he dwelleth with you, and shall be in you."),
    ('p', "Now I've told you many times, the language of the New Testament is what? IN. IN. Someone say IN."),
    ('p', "So we said, you know we had two prophecies, remember that? Ezekiel's prophecy and Joel's prophecy. We said which one is closer to the New Testament? Ezekiel, because Ezekiel says IN. Joel says UPON. But Ezekiel says IN."),
    ('p', "IN is the language of the New Testament: in Christ, in you. Praise the Lord! Praise the Lord! Come on guys, are you here? Praise the Lord!"),
    ('p', "Remember, Paul never says to Philemon acknowledge every good thing that is upon you. Is that what he says? That is where? In you. In you in Christ. That's what you're supposed to find out."),
    ('p', "So he says in you a well, you're born of the Spirit, he gave you his Spirit, right, because he will abide in you forever."),

    # Luke 24 & Acts 1
    ('p', "Now we have two other statements: Luke's Gospel 24:48, 49. 49, Luke's Gospel 24:49, Luke says, Luke's Gospel 24:49, Luke says:"),
    ('ref', "Luke 24:49"),
    ('v', "[49] And, behold, I send the promise of my Father upon you: but tarry ye in the city of Jerusalem, until ye be endued with power from on high."),
    ('p', "Then Acts 1:8, Acts 1:8, he says:"),
    ('ref', "Acts 1:8"),
    ('v', "[8] But ye shall receive power, after that the Holy Ghost is come upon you: and ye shall be witnesses unto me both in Jerusalem, and in all Judaea, and in Samaria, and unto the uttermost part of the earth."),
    ('p', "And these are Jesus's words. And so that could get us a little bit confused, because we have in you, we have upon you. Is he talking about two different spirits or two different experiences of the same spirit?"),
    ('p', "Now I will advise you to stay close to the second service where I'm teaching on the power and revelation, and I'm going to explain a little bit of that in the language for and what it means."),
    ('p', "Because sometimes you know in the study there's some traditional beliefs that have persisted for many years in the church world because the people who started them were very powerful leaders."),
    ('p', "And if someone is very powerful and he starts a doctrine, because many folks will hardly find out whether it's so or not, it can take a long time, a very long time."),
    ('p', "Just like I told you the doctrine of rapture, the way it's said today. Somebody started it and then everybody just followed the train and nobody really asked the right questions. Praise the Lord! And quite a number of things like that. But you need to know why."),
    ('p', "So when we say the Spirit upon you and the Spirit within you, and we can say well, the same Spirit but two experiences, but really not so."),
    ('p', "We'll find out that setting in the second service when we're looking at the power and revelation, because that's what sorts it out for us."),

    # Romans 8
    ('p', "Now, are you listening? Okay, good. Now so let's just take our mind a little bit away from how Luke describes that, because that may take us off the mark a little bit."),
    ('p', "Now look at Paul in Romans 8. Paul is talking about the Spirit, right? In Romans 8, he says:"),
    ('ref', "Romans 8:1-2"),
    ('v', "[1] There is therefore now no condemnation to them which are in Christ Jesus, who walk not after the flesh, but after the Spirit."),
    ('v', "[2] For the law of the Spirit of life in Christ Jesus hath made me free from the law of sin and death."),
    ('p', "'The law of the Spirit of life where? In.' Notice Paul's language for the Spirit is IN. 'The law of the Spirit of life in,' someone say IN, 'in Christ Jesus.'"),
    ('p', "So the Spirit of life is where? In Christ Jesus. Now look at verse 8, are you there? Now again, this is not the thesis, but just pay attention. Again he mentions in, right, in verse 4, 'fulfilled in us.' Are you there? Are you there? Fulfilled in us."),
    ('ref', "Romans 8:4"),
    ('v', "[4] That the righteousness of the law might be fulfilled in us, who walk not after the flesh, but after the Spirit."),
    ('p', "So verse 8 says, 'those that are in the flesh, they that are in the flesh cannot please God.'"),
    ('p', "Now let's take 9 together. No, hold on again, don't forget: Romans 8:1, 'There is therefore now no condemnation to them which are where? In Christ.'"),
    ('p', "Now verse 2, 'The law of the Spirit of life what? In Christ.' Very vital words: in Christ, in Christ. Verse 4, pardon me, 'fulfilled where? In us.'"),
    ('p', "So look at verse 9, let's go: 'But ye are not in the flesh, but where? In the Spirit.'"),
    ('ref', "Romans 8:8-11"),
    ('v', "[8] So then they that are in the flesh cannot please God."),
    ('v', "[9] But ye are not in the flesh, but in the Spirit, if so be that the Spirit of God dwell in you. Now if any man have not the Spirit of Christ, he is none of his."),
    ('v', "[10] And if Christ be in you, the body is dead because of sin; but the Spirit is life because of righteousness."),
    ('v', "[11] But if the Spirit of him that raised up Jesus from the dead dwell in you, he that raised up Christ from the dead shall also quicken your mortal bodies by his Spirit that dwelleth in you."),
    ('p', "Now listen very well, study, you need to study it now: he says you are where? In Christ. He says the law of the Spirit of life where? In Christ. There is therefore now no condemnation to them who are where? In Christ. The righteousness of the law is fulfilled where? In us."),
    ('p', "Listen: in Christ, in Christ Jesus, in us, then he says you are not where? In the flesh, but where? In the Spirit."),
    ('p', "Now listen to that explanation: 'if so be that, uh, the Spirit of God what?' So the word in Christ and the indwelling of the Spirit, are they synonymous? Uh? Does the word in Christ mean the Spirit of God dwells in you? Come on, right? Come on, are you there?"),
    ('p', "'Now if any man have not who? The Spirit of Christ, he is none of his.' Look at that: he's none of his. Now look at 9, 10 pardon me, 'And if Christ be where? Uh? In you.'"),
    ('p', "So that means in Christ and Christ in you are synonymous terms. If you are in Christ, it means Christ is in you. Are you there? Are you listening? Do you understand what I just did for you now?"),
    ('p', "Okay, so when we say in Christ and Christ in you, it's the same truth that is being explained differently."),
    ('p', "He says, 'And if Christ be in you, the body is dead because of sin; but the Spirit is life because of righteousness. But if the Spirit of him that raised up Jesus from the dead dwell where? In.'"),
    ('p', "Now look at how many times the word in you, in you, in is used again: Romans 8:1, Romans 8:2, Romans 8:4, Romans 8:9, Romans 8:10, Romans 8:11, where? In. Hallelujah! In you, in Christ."),

    # Romans 8:12-16
    ('p', "Now go further. He says, 'he that raised up Christ from the dead shall also quicken your mortal bodies by his Spirit that dwelleth where? In you.'"),
    ('p', "Look at 12: 'Therefore, brethren, we are debtors, not to the flesh, to live after the flesh,' that's a fact."),
    ('p', "'For if you,' now if you, conditional statement, right? 'If you, if you live after the flesh, you shall die.' You see, that statement, is it a possibility or an impossibility? I thought you know Bible interpretation, what is it? Uh?"),
    ('p', "He says, 'brethren, can a man who is called a brother in Christ live after the flesh, eh? No.'"),
    ('p', "So the word if there is telling you, is painting a possibility that's an impossibility with the believer in Christ. Because he says to you, he says, 'if you live after the flesh, you shall die: but if you through the Spirit do mortify the deeds of the body, you shall live.'"),
    ('p', "I explained that word live to you before, right? Uh? Live different from walk, can you remember that? Good."),
    ('p', "So verse 14: 'For as many as are led,' in other words, verse 14 explains why you're not in verse 13, 'For as many as are led by the Spirit of God, they are the sons of God, okay?"),
    ('p', "Verse 15 tells you what led by the Spirit of God means. He says, 'For ye have not received the spirit of bondage again to fear; but ye have received the Spirit of adoption, whereby we cry, Abba, Father. The Spirit itself beareth witness where? With our spirit, that we are the children of God.'"),
    ('ref', "Romans 8:12-16"),
    ('v', "[12] Therefore, brethren, we are debtors, not to the flesh, to live after the flesh."),
    ('v', "[13] For if ye live after the flesh, ye shall die: but if ye through the Spirit do mortify the deeds of the body, ye shall live."),
    ('v', "[14] For as many as are led by the Spirit of God, they are the sons of God."),
    ('v', "[15] For ye have not received the spirit of bondage again to fear; but ye have received the Spirit of adoption, whereby we cry, Abba, Father."),
    ('v', "[16] The Spirit itself beareth witness with our spirit, that we are the children of God:"),
    ('p', "Are you still there? Are you still there? Someone say IN. Okay, so that's the word IN."),
    ('p', "So John's rendering, because you know Jesus spoke in different types of revelation to them, John's rendering is the closest to the New Testament, when he says he will dwell in you. That's the New Testament. He will dwell in you."),
    ('p', "So Luke's explanation of upon is valid, but it's a part of the whole truth. Are you listening? Are you listening? Okay, let's move forward a bit."),

    # Galatians 4
    ('p', "Now look at Galatians 4. Verse 4:"),
    ('ref', "Galatians 4:4-6"),
    ('v', "[4] But when the fulness of the time was come, God sent forth his Son, made of a woman, made under the law,"),
    ('v', "[5] To redeem them that were under the law, that we might receive the adoption of sons."),
    ('v', "[6] And because ye are sons, God hath sent forth the Spirit of his Son into your hearts, crying, Abba, Father."),
    ('p', "'But when the fulness of the time was come, God sent forth his Son, made of a woman, made under the law, to redeem them that were under the law, that we might receive the adoption of sons.'"),
    ('p', "Okay, let's take verse 6 together: 'And because ye are sons, God hath sent forth the Spirit of his Son into your hearts, crying, Abba, Father.'"),
    ('p', "The Spirit of his Son into your hearts, crying, Abba, Father."),
    ('p', "1 Corinthians 2. Praise the Lord!"),
    ('p', "So salvation is to receive the Spirit. Is that clear? Is that clear?"),
    ('p', "Salvation is not just that you were transferred from darkness to light and then you can go back again, no. Salvation means you have received the Spirit of God. That's what it is, and you are in him."),

    # 1 Corinthians 2
    ('p', "Now 1 Corinthians 2 verse 9:"),
    ('ref', "1 Corinthians 2:9-12"),
    ('v', "[9] But as it is written, Eye hath not seen, nor ear heard, neither have entered into the heart of man, the things which God hath prepared for them that love him."),
    ('v', "[10] But God hath revealed them unto us by his Spirit: for the Spirit searcheth all things, yea, the deep things of God."),
    ('v', "[11] For what man knoweth the things of a man, save the spirit of man which is in him? even so the things of God knoweth no man, but the Spirit of God."),
    ('v', "[12] Now we have received, not the spirit of the world, but the spirit which is of God; that we might know the things that are freely given to us of God."),
    ('p', "'But as it is written, Eye hath not seen, nor ear heard, neither have entered into the heart of man, the things which God hath prepared for them that love him. But God hath revealed them unto us by his Spirit: for the Spirit searcheth all things, yea, the deep things of God. For what man knoweth the things of a man, save the spirit of man which is in him? even so the things of God knoweth no man, but the Spirit of God.'"),
    ('p', "I want us to read verse 12 together, let's go: 'Now we have received,' hold on, hold on. You know the first two verses were to explain something to you, okay? Now the verse 12 explains salvation."),
    ('p', "The first two verses do not explain salvation to you; the first two verses simply tell you, so God has revealed those things to us by his Spirit, the Spirit searcheth all things, the deep things of God, he says the things of God knows no man. He's talking about the Spirit there, he's not talking about salvation."),
    ('p', "So in verse 12, verse 12 is why he explained verse 10 and 11. So let's take verse 12 again, let's go, come on: 'Now we have received,' hold on, he didn't say we will receive, uh? Come on, did he say we will receive?"),
    ('p', "We have received: past tense, present tense, future tense? Uh? Which tense? So you have received. Say I have received! Past tense. Now, I have received what? Uh?"),
    ('p', "'Not the spirit of the world, but the spirit which is of God; that we might know the things that are freely given to us of God.'"),
    ('p', "Hallelujah! Hallelujah! In other words, your spirit knows. Your spirit already knows, that you might know. He didn't say that that we'll know, so you might know the things that are freely given to us of God."),
    ('p', "So salvation is to receive the Spirit of God. Hallelujah! So what is salvation? It's to receive the Spirit of God."),
    ('p', "So you cannot be saved and not have the Spirit of God, because salvation is to receive the Spirit of God. Hallelujah! Are you listening to that?"),
    ('p', "So if you receive the Spirit of God, you don't, I had a theory that says when you're born again, okay, you are now in Christ, that's how they say it."),
    ('p', "But he said, okay let me put it the way he said it: he said you can get the dirt, you can take the man out of the dirt, so when you come out of the dirt, now I heard it from a very respectable preacher that I respected, so you take the man out of the dirt, that is salvation; so but it's time to take the dirt out of the man, that's sanctification and deliverance. You see what they say?"),
    ('p', "So you will need, they believe that you still have some very foreign spirits that lie somewhere in you, and so someone will feel you have a dual personality disorder."),
    ('p', "That is, you are in Christ, but not totally for Christ; that you are born again, but you have some other spirits that, you know, whether tenants or landlords, who walk together in your house, in your body, and then you, then sometimes your experiences don't help you."),

    # Anecdote: Sickness & Water Goddess Deliverance
    ('p', "When I got saved, I used to fall sick. So I went to someone to counsel with me, so he asked me some questions."),
    ('p', "And he said I should give him some stories about my family, so I did give him some stories. Sometimes you help those kind of people to give you more unbelief."),
    ('p', "So I told him that well, my family, my father told me this story, my uncle told me, he said, ah! He said, it is those spirits that are coming after you."),
    ('p', "Because some people believe that you can go to a water goddess and get children. And so they say, well, what were your parents doing before you were born?"),
    ('p', "So well, they used to worship at the shrine. Ow! The guy shakes his head. You see, you're in trouble. So what do you have to say? Well, they asked for you from the goddess. So the goddess gave you to your parents. So what's going to happen now? You have to go and beg them. Or you tell them to take their hands off you, something like that."),
    ('p', "Some people believe that kind of stuff, and then they say go and ask your mother what happened when you were born."),
    ('p', "So that makes you home, you get near to mom, say mommy tell me the truth about how I was born. You know, and when you are an African, you know, it takes a lot of scriptural sense away from your mind, it's just really spooky."),
    ('p', "You know, but in addressing this, I tell folks before, I said Abraham came from an idol-worshipping home. It's not likely his father prayed to God. You know that true?"),
    ('p', "You know like, if your father was an idol worshipper, it's not likely to mention, he didn't even know who God, can't call him Jehovah, can't have called him El Shaddai, all those things. He just probably talked to the goddess or god he worshipped."),
    ('p', "And God didn't say, Abraham I wanted to talk to you, but you need deliverance."),
    ('p', "The point is, everybody who's born into this world, from Paul's theology, the Bible says sin passed upon all men. So either way, you came from a goddess, you know what I mean anyway. I think it's just a wrong theology. A goddess can't give anybody a child, because they have no such capacity. Is that clear?"),
    ('p', "Some of you are happy because you heard that, because you thought that you came from, you came from under the rock. You know, so sometimes your experiences deceive you."),
    ('p', "And I know folks today that they are doubting their salvation because of certain habits they have."),
    ('p', "Yes, you should walk in the word and break those habits, but you see, your salvation is not a product of your experiences. Your salvation is a gift from God. It's not a product of your experiences. Your experiences sometimes can be good, sometimes can be bad, sometimes they are just neutral. But salvation is a product of God's gift to you. Is that clear? Is that clear? Come on guys, is that clear? All right, good."),
    ('p', "So he gave us his Spirit, so we are not, we're not, we're not, we are already, and I say it often, that it's the Spirit that God gave you that makes you holy."),
    ('p', "Some guys will say, well, the Holy Spirit can never come into a sinful, a sinful, a sinful vessel. So what does a sinful vessel do? Say the vessel has to become a righteous vessel first, a holy vessel, so the Holy Spirit comes here to do what? He's scared of what is not holy? No, no."),
    ('p', "It's the Holy Spirit that comes into the life of a sinner and makes the sinner holy. God only gives his Spirit to sinners that believe. Amen! Is that clear? When you believe, he gives his Spirit to sinners. So that's the word of God."),
    ('p', "At salvation, you receive the Holy Spirit. Is that clear? Because that's what salvation is. Praise the name of Jesus. Is that clear?"),

    # Ephesians 1 & 4 - Sealed
    ('p', "Now look at Ephesians 1. Ephesians 1:13. It says:"),
    ('ref', "Ephesians 1:13"),
    ('v', "[13] In whom ye also trusted, after that ye heard the word of truth, the gospel of your salvation: in whom also after that ye believed, ye were sealed with that holy Spirit of promise,"),
    ('p', "Same thing in chapter 4 and verse 30:"),
    ('ref', "Ephesians 4:30"),
    ('v', "[30] And grieve not the holy Spirit of God, whereby ye are sealed unto the day of redemption."),
    ('p', "That word sealed is used in the Roman Empire for that which is a proof of ownership, a proof of permanent ownership. So he says the Holy Spirit is a proof of permanent ownership. Are you there? Are you listening? It's a proof of permanent ownership."),
    ('p', "So in salvation, God claims the sinner totally. He claims his past, he claims his present, and he claims his future. That's why it's the giving of the Spirit."),
    ('p', "You don't, you don't say, well, God didn't make us halfway, you came halfway, you came halfway. He made us through, through the way down into the future."),
    ('p', "He has sealed us with his Spirit. Say I'm sealed with the Spirit! Now where is the Spirit? In you. So that's a seal, a proof of permanent ownership."),

    # Spiritual DNA & Permanent Ownership
    ('p', "Now let me ask you quickly. You know we said Jesus used three examples in John 3, John 4, John 7. He said receive. Now in John 4:14, he says everlasting what? Life. In John 3, he says born of the Spirit. In John 7, he says receive. How about John 16:12 and 13? He says, so he tells you it will dwell in you forever. I want you to listen to this. And we said being a seal is a proof of permanent ownership."),
    ('p', "So Jesus says you are born of the Spirit, he says that which is born of the Spirit is what? Spirit. So in other words, listen to this: the Spirit of God is the DNA of the Christian. That's his heredity. That's why it's a permanent ownership. It's not, you didn't receive something that you will, you can give up."),
    ('p', "It's like a birth. You can never, you can disown someone in the court because you know he got you angry, your son and all that, but you can never disclaim the fact that the child, the boy or the girl carries your life, or your DNA, or your genes. That can never be altered. Is that true? Is that true? Can never be altered."),
    ('p', "So the man who has received the gospel carries in him and with him the very DNA of God, the Spirit of God. So the Spirit of God is not an experience. The Spirit of God is the life that we have from God. Are you listening? That's the life that we have from God."),
    ('p', "The way a child receives life from the parent, the life we receive from God is the Spirit of God."),
    ('p', "Hallelujah! Let me see your hand if you understand that, let me see your hand, come on wave your hand. Is that clear to you? So it's not an experience; the Spirit of God is a life that we have of God."),

    # John 16 - Intro into Truth
    ('p', "So John 16, however, you know in all of Jesus's teachings, we only find an intro into the truth, we don't find an explanation. In all of this. So you know we're discussing, don't forget, the gifts of the Spirit, and Jesus never mentioned the word gifts of the Spirit."),
    ('p', "He mentioned the word gift in John 4:10, so you saw that, but he doesn't mention gifts of the Spirit, just like Jesus never mentioned grace. Hallelujah! Hallelujah! Jesus never said, Jesus never said in Christ. Does he say in Christ? He didn't mention grace."),
    ('p', "But look at John 16, he says there are many things that I desire to say to you, John 16:12:"),
    ('ref', "John 16:12-13"),
    ('v', "[12] I have yet many things to say unto you, but ye cannot bear them now."),
    ('v', "[13] Howbeit when he, the Spirit of truth, is come, he will guide you into all truth: for he shall not speak of himself; but whatsoever he shall hear, that shall he speak: and he will shew you things to come."),
    ('p', "So Jesus is telling you that in his teaching ministry, he gives an intro. You get that? He gives an intro, he doesn't give you the full picture."),
    ('p', "So if you are discussing gifts of the Spirit, we're not going to get it in the four Gospels, but we can get an intro in the four Gospels. Praise the Lord! Are you there?"),

    # John 14 - Greater Works
    ('p', "Let's see an idea of where he begins to talk about the gifts of the Spirit. Again, salvation is what? Uh? You know you can call salvation forgiveness of sins, right? We can also call salvation what? Receiving the Holy Spirit, or the gift of the Spirit. Now don't forget, not gifts in plural, but just singular, right? The gift of the Spirit is salvation."),
    ('p', "Now look at John 14 verse 12. John 14 verse 12. Are you there? Come on, are you there? What does he say? Uh? Verily, verily, that word comes up a lot. Verily, verily, I say unto you. Who is verily verily? Nobody called verily here? Verily, verily, I say unto you, let's go:"),
    ('ref', "John 14:12-14"),
    ('v', "[12] Verily, verily, I say unto you, He that believeth on me, the works that I do shall he do also; and greater works than these shall he do; because I go unto my Father."),
    ('v', "[13] And whatsoever ye shall ask in my name, that will I do, that the Father may be glorified in the Son."),
    ('v', "[14] If ye shall ask any thing in my name, I will do it."),
    ('p', "'He that believeth on me,' hold on. Now he that believeth on him, what happens to he that believes? Uh? Uh-uh, please stop there. In the Book of John, he that believes, what happens to he that believes? He receives. Is that clear? Uh? Come on guys, is that clear?"),
    ('p', "He that believes, in John 7:37 to 39, will receive the Spirit; in John 4:14, he that believes will receive the Spirit, right? Come on, isn't that right?"),
    ('p', "So can we say, verily, verily, I say unto you, he that receives my Spirit, would that be correct? Is that correct? Okay, so go further: he that receives my Spirit, go further, 'the works that I do shall he do also, right? The works that I do shall he do also; and greater works than these; because I go unto my Father.'"),
    ('p', "And there's some very horrific interpretations of this verse. I know I've explained this many times to you, and the fellow said, what are the greater works? So what is greater? You hear what he says, he says, Jesus was saying greater works, he's trying to explain, he said Jesus worked miracles, he raised the dead, he healed the sick; Jesus is saying here, now this is how people they color their own imagination, he said, it's a theology, respected theology, he is a, he's a, he pastored the Metropolitan Tabernacle, the one started by Charles Spurgeon."),
    ('p', "And so he said, now I'm saying this because I've written him already, so I'm not backbiting. I can't be backbiting on CD anyway, that is media-biting. So he said Jesus is saying here that the days of miracles will be over."),
    ('p', "And I found that interesting, because a greater work, what is greater than what Jesus did, that's what he said. Salvation of souls is a greater work."),
    ('p', "And then someone who read the book, he's a pastor, he said I never saw it like that, and I said neither did I. Because I'm sure when you said, by God, revelation! I said if Jesus said a greater work, now I understand. He said greater works. How many salvations? Can you hear that? And I said oh, I never saw it like that, I said certainly, you always don't see anything."),
    ('p', "And I said, and when I wrote him I said this, I said Jesus didn't say he that believes in me, the works I do he won't be able to do anymore. I said sir, even if I agree your greater work is salvation, he said the works I do he will first be doing, and greater works than this. So by the way sir, are we doing salvation? Amen. You decide who comes sir."),
    ('p', "But the work there is ergon in the Greek, deeds, all right, a definite act, ergon, where you have energeo, words like that."),
    ('p', "So he says greater works than this, in other words he's pointing to a particular work: the works I do he will do also. Look at that word also, also! That means he'll raise the dead, he'll heal the sick, he'll cast out demons, and hallelujah, greater works than this, he didn't say a greater work, because I go unto my Father. Hallelujah! Glory to God!"),
    ('p', "So you can see Jesus is speaking about the gifts. Now how do you know? Don't forget, the language of the New Testament is what? IN."),
    ('p', "So read the next verse, let's take those verses together, next verse: 'And whatsoever ye shall ask in my name, that will I do, that the Father may be glorified where? In the Son. If ye shall ask any thing in my name, I will do it. If ye love me, keep my commandments.'"),
    ('p', "Don't forget, he didn't say keep the Ten Commandments. Is that clear? He didn't say keep the two commandments, new commandment. What's the commandments? He's giving them instructions, right? Keep my instructions. What instructions? What do our people say, talk about Jesus's commandments? What are they? Thou shall not worship any other god but me, is that what his commandments are?"),
    ('p', "Verse 16:"),
    ('ref', "John 14:15-20"),
    ('v', "[15] If ye love me, keep my commandments."),
    ('v', "[16] And I will pray the Father, and he shall give you another Comforter, that he may abide with you for ever;"),
    ('v', "[17] Even the Spirit of truth; whom the world cannot receive, because it seeth him not, neither knoweth him: but ye know him; for he dwelleth with you, and shall be in you."),
    ('v', "[18] I will not leave you comfortless: I will come to you."),
    ('v', "[19] Yet a little while, and the world seeth me no more; but ye see me: because I live, ye shall live also."),
    ('v', "[20] At that day ye shall know that I am in my Father, and ye in me, and I in you."),
    ('p', "'And I will pray the Father, and he shall give you what?' See, that's where he started from, can you see it? 'He shall give you another Comforter, that he may abide with you forever; Even the Spirit of truth; whom the world cannot receive, because it seeth him not, neither knoweth him: but ye know him; for he dwelleth with you, and shall be where? In you. I will not leave you comfortless: I will come to you. Yet a little while, and the world seeth me no more; but ye see me: because I live, ye shall live also. At that day ye shall know that I am where? In my Father, and ye in me, and I in you.'"),
    ('p', "So that's salvation. Can you see that? So in John 14, he didn't use the word gift, but he explains the working of the gifts, the workings of the supernatural. The believer can do it because he has received the Spirit, and he's going to work through him. Hallelujah!"),

    # Mark 16 & Gift of the Spirit in Acts
    ('p', "Now look at Mark 16. In Mark 16, he says:"),
    ('ref', "Mark 16:17-18"),
    ('v', "[17] And these signs shall follow them that believe; In my name shall they cast out devils; they shall speak with new tongues;"),
    ('v', "[18] They shall take up serpents; and if they drink any deadly thing, it shall not hurt them; they shall lay hands on the sick, and they shall recover."),
    ('p', "Signs that follow, and he calls, Jesus calls them signs: cast out demons, speak in tongues, lay hands on the sick. So he calls them signs. Did he call them gifts? No, he didn't call them gifts. He calls them signs."),
    ('p', "Now in the Book of Acts, we find the word gifts of the Spirit used for what? Now don't forget, in Acts 2:38, Peter said:"),
    ('ref', "Acts 2:38"),
    ('v', "[38] Then Peter said unto them, Repent, and be baptized every one of you in the name of Jesus Christ for the remission of sins, and ye shall receive the gift of the Holy Ghost."),
    ('p', "What is the gift of the Holy Ghost there? Salvation."),
    ('p', "Acts 8:20, Simon the sorcerer offered money, and Peter said:"),
    ('ref', "Acts 8:20"),
    ('v', "[20] But Peter said unto him, Thy money perish with thee, because thou hast thought that the gift of God may be purchased with money."),
    ('p', "That gift of God there is salvation, right?"),
    ('p', "So in the Book of Acts, the word gift of the Spirit is used for salvation. Don't forget that, the gift of the Spirit is used for salvation."),

    # Utterance Gifts in Acts
    ('p', "Now, what did the Book of Acts call the manifestations of the Spirit? Are you listening? You need to know this, so that when you're reading Acts, you're not confused."),
    ('p', "In Acts 2:16, Peter said:"),
    ('ref', "Acts 2:16-17"),
    ('v', "[16] But this is that which was spoken by the prophet Joel;"),
    ('v', "[17] And it shall come to pass in the last days, saith God, I will pour out of my Spirit upon all flesh: and your sons and your daughters shall prophesy, and your young men shall see visions, and your old men shall dream dreams:"),
    ('p', "Now notice what he used: pour out of my Spirit upon all flesh."),
    ('p', "Acts 8:15 and 16, Peter and John went to Samaria, prayed for them, that they might receive the Holy Ghost:"),
    ('ref', "Acts 8:15-16"),
    ('v', "[15] Who, when they were come down, prayed for them, that they might receive the Holy Ghost:"),
    ('v', "[16] (For as yet he was fallen upon none of them: only they were baptized in the name of the Lord Jesus.)"),
    ('p', "What word did he use? Fallen upon."),
    ('p', "Acts 10:44, Cornelius's house:"),
    ('ref', "Acts 10:44-46"),
    ('v', "[44] While Peter yet spake these words, the Holy Ghost fell on all them which heard the word."),
    ('v', "[45] And they of the circumcision which believed were astonished, as many as came with Peter, because that on the Gentiles also was poured out the gift of the Holy Ghost."),
    ('v', "[46] For they heard them speak with tongues, and magnify God."),
    ('p', "Poured out, fell upon."),
    ('p', "Acts 11:15, Peter explaining:"),
    ('ref', "Acts 11:15, 17"),
    ('v', "[15] And as I began to speak, the Holy Ghost fell on them, as on us at the beginning."),
    ('v', "[17] Forasmuch then as God gave them the like gift as he did unto us, who believed on the Lord Jesus Christ; what was I, that I could withstand God?"),
    ('p', "Acts 19:6, Paul at Ephesus:"),
    ('ref', "Acts 19:6"),
    ('v', "[6] And when Paul had laid his hands upon them, the Holy Ghost came on them; and they spake with tongues, and prophesied."),
    ('p', "Came on them."),
    ('p', "So the Book of Acts uses: poured out, fell upon, came upon, received the gift. For what? Speaking with tongues, prophesying, praise the Lord! That is, the utterance gifts!"),
    ('p', "In the Book of Acts, the utterance gifts are described as the Spirit coming upon, falling upon, poured out. Are you listening? Are you listening? Good."),

    # Revelation Gifts in Acts
    ('p', "Now what about revelation gifts? Word of knowledge, word of wisdom, discerning of spirits. How does the Book of Acts describe that? Does it say the Spirit fell upon them? No!"),
    ('p', "Look at Acts 5:3, Peter said to Ananias, 'why hath Satan filled thine heart to lie to the Holy Ghost?' Verse 9, 'How is it that ye have agreed together to tempt the Spirit of the Lord?' What word is used? The Holy Ghost, the Spirit of the Lord."),
    ('ref', "Acts 5:3-4, 9"),
    ('v', "[3] But Peter said, Ananias, why hath Satan filled thine heart to lie to the Holy Ghost, and to keep back part of the price of the land?"),
    ('v', "[4] Whiles it remained, was it not thine own? and after it was sold, was it not in thine own power? why hast thou conceived this thing in thine heart? thou hast not lied unto men, but unto God."),
    ('v', "[9] Then Peter said unto her, How is it that ye have agreed together to tempt the Spirit of the Lord? behold, the feet of them which have buried thy husband are at the door, and shall carry thee out."),
    ('p', "Look at Acts 10. Peter is on the housetop praying, he falls into a trance, sees a sheet coming down with all manner of animals, four-footed beasts, creeping things, fowls of the air."),
    ('p', "And he was told, Peter, rise, kill and eat. And he says, no. Now did Peter say no unconsciously? No. You normally reflect your spiritual status in your dreams and visions. That's what you reflect."),
    ('p', "Some of you have bad dreams a lot, and they come to pass. I know somebody that never has good dreams, never! I've never heard the person say you are blessed from the dream. It's either monkeys or dogs. When you start having dreams like that, you are living in fear. Because the Old Testament tells you that out of the multitude of business ariseth a dream. Are you listening? You wake up and you say, you're sweating, Jesus, Jesus, you know that's how you are."),
    ('p', "Peter is saying no. Jesus took 40 days to explain the word to them, he comes back in a vision, Peter is still saying no. No, no, three times! I mean, even the psalmist says, once has the Lord spoken, twice have I heard, and three times the guy is saying no."),
    ('p', "Look at one thing: when he got up, the Bible says in verse 17, Peter doubted in himself what this, you know he's doubting in the vision. Hallelujah! I think I explained that to you, right? Lazarus, rich man, on Wednesday? Uh? Okay, let's move forward."),
    ('p', "So in verse 19:"),
    ('ref', "Acts 10:19"),
    ('v', "[19] While Peter thought on the vision, the Spirit said unto him, Behold, three men seek thee."),
    ('p', "So that the Spirit said to him, that's a revelation, right? Right? So come on, is that right? So two times. Now in Acts 10, sorry Acts 5:4 and 9, it's the Spirit. Now again, the Spirit said. Notice what's used for revelation is THE SPIRIT. Come on guys, are you following? Hey, are you following this? Are you following this?"),
    ('p', "Look at chapter 11: Agabus, Agabus, Agabus, Agabus. Agabus, the unconscious grandfather of the charismatic movement:"),
    ('ref', "Acts 11:28"),
    ('v', "[28] And there stood up one of them named Agabus, and signified by the Spirit that there should be great dearth throughout all the world: which came to pass in the days of Claudius Caesar."),
    ('p', "Now how would he show them by the Spirit? Did he say, okay, Spirit becomes a projector machine, and you know, he must have done what, said it by speaking, right? Come on guys, uh? Hey, are you listening? By speaking. So he showed them by the Spirit. Now is that revelation? So are you seeing the word the Spirit, by the Spirit, the Spirit said, is used for what? Thank you."),
    ('p', "Okay, go to Acts 13. Are you there?"),
    ('ref', "Acts 13:2"),
    ('v', "[2] As they ministered to the Lord, and fasted, the Holy Ghost said, Separate me Barnabas and Saul for the work whereunto I have called them."),
    ('p', "Two things: future and past, revelation. The Holy Ghost said. Are you listening? So the Holy Ghost said, or the Spirit said, or they showed by the Spirit, refers to what? Revelation. And Peter doesn't, there's no word like gift for that one: the Holy Ghost said. But utterance is the Holy Ghost upon. Are you listening? Are you following it? All right."),

    # Acts 20 & 21 - Agabus & Utterance carrying Revelation
    ('p', "Look at Acts 20 verse 22. You know with these terms, I'm clarifying it for you so that when you're studying them, you understand it. Is that clear? Acts 20:22:"),
    ('ref', "Acts 20:22-23"),
    ('v', "[22] And now, behold, I go bound in the spirit unto Jerusalem, not knowing the things that shall befall me there:"),
    ('v', "[23] Save that the Holy Ghost witnesseth in every city, saying that bonds and afflictions abide me."),
    ('p', "Now the Holy Ghost said, now the Holy Ghost said here will be what again? Uh? Revelation."),
    ('p', "Now go to 21, I hope you are paying attention. 21 verse 4:"),
    ('ref', "Acts 21:4"),
    ('v', "[4] And finding disciples, we tarried there seven days: who said to Paul through the Spirit, that he should not go up to Jerusalem."),
    ('p', "This is a very clear text."),
    ('p', "Now look at verse 11, that's Agabus again:"),
    ('ref', "Acts 21:11"),
    ('v', "[11] And when he was come unto us, he took Paul's girdle, and bound his own hands and feet, and said, Thus saith the Holy Ghost, So shall the Jews at Jerusalem bind the man that owneth this girdle, and shall deliver him into the hands of the Gentiles."),
    ('p', "So the Holy Ghost said is revelation in the Book of Acts, correct? But how do we receive that revelation? By what? Utterance. It appears that the church in Acts graduated from the utterance gifts to, not graduated, increased from the utterance gifts to the revelation gifts. So that means within the revelation gifts, all right, you will have the utterance gifts."),
    ('p', "Now did you observe something? Okay, whilst Peter called the utterance gifts the gift, doron, there was no particular name or explanation for the, for Paul pardon me, for the revelation gifts. However, if you don't get this one, you should pray not just a four-line prayer, say Isaiah's prayers: that means a man who walks in the utterance gifts can equally walk in what? Revelation! Because it's about speaking. You get that?"),
    ('p', "Agabus spoke out revelation; the guys in Caesarea, they spoke out revelation; Agabus spoke out revelation; Acts 13:2, they spoke out by what? Revelation. All right, so the man who is into the utterance gifts can also speak revelation gifts. Praise the Lord! Is that clear? Is that clear? Okay."),

    # Anecdote: Sister Joseph, Chucky, Pastoring the Sheep
    ('p', "Now what about the power gifts? Miracles, signs, and wonders. You'll see that in the Book of Acts: miracles, signs, and wonders."),
    ('p', "So you may have the idea that there are different, maybe you call it, this guy has an anointing for, you know that's how they say it. They say, that man is, have you heard this, hear this very well: 'He host Reverend Somi with other anointed men of God.' Can you have a man of God who is not anointed? Well, in that context, yeah. Other anointed men of God. 'Ah that brother, he's anointed oh! When he's speaking like this, he speaks the very oracles of God. That sister, she sees things.'"),
    ('p', "There was one sister in our fellowship then, every fellowship time she has dreams. She's like Sister Joseph. Sister Joseph. You know Joseph had dreams, she had many many dreams, very very bad ones. And she'd be quiet: 'I sense that somebody here is in danger. I just sense you just have to pray.'"),
    ('p', "That's why if you're a pastor, don't allow people to just come forward to give a word. If you're a pastor, you ask, what's the word?"),
    ('p', "You know where I learnt it from? There was this ministry headed by a friend of mine, now with the Lord now, we call him Chucky. So they used to have Holy Ghost meetings, so I used to go there, University of Ibadan, and yeah people give words, exhortation, prophecy, word of wisdom, word of knowledge. So they will write it down because pastor said they'll read it out."),
    ('p', "So one day, a young sister, and I know this sister very very well, she's still my friend, so they will pass the paper across. Everybody was you know having a great time, laughing in the Spirit, having a great time, ministering to folks."),
    ('p', "So he just said, he wasn't paying attention, he just said, praise the Lord, yeah this is a word: 'All these things that you are doing, there are people speaking in tongues who are in hell right now! Keep your salvation.' I was watching and I know, okay, that's headless. Don't read things without seeing them first. Praise the Lord!"),
    ('p', "Sometimes some people get very offended. They say, pastor, that was the word for the church, take time, speak this for the church. They say, I will tell them. You tell me first. And if you want to say daddy, say it here, I am the pastor here. Pastor is the door of the sheep, all right?"),
    ('p', "No, you don't go behind to enter a man's house, so let me be the one to allow you to enter or not to enter. That's why you press bells in houses. You don't just open the gate, 'Noah!' Say I'm here, can I come in? And you're already in. So you don't do that. Praise the Lord!"),
    ('p', "Even testimony, sometimes tell them what's the testimony. Let me tell you, church members get hurt. Somebody will give a testimony and say things like, 'Ah, let's be careful for the devil!' Oh my god! And they overcame the church by the word of the Satan's testimony! Praise the Lord! You learning something? All right."),

    # 1 Corinthians 12:1 - Pneumatikos & Charisma
    ('p', "So that means the Book of Acts doesn't really clarify how those things happen. We can see they spoke it, we can see they heard things, we can see they, but we don't know how. Spirit upon."),
    ('p', "That's why it got many of us to think that you know I was in this meeting and the Spirit of God just came upon me, you see that? Just came upon me, I didn't know what I was saying again. I was just seeing some things, I was just seeing some things. In fact, the kind of tongues I was speaking was strange, it was high tempo, like machine gun! I was just seeing some things that I didn't know that I was seeing these things, you know, I was just seeing some things. As I wanted to continue, the Spirit had gone! Hallelujah!"),
    ('p', "It's because you read the Book of Acts."),
    ('p', "1 Corinthians 12:1. And that's my teaching for today. Amen. For the last few weeks, I've just begun to have what I wanted to say, but you get blessed anyway, uh? All right."),
    ('ref', "1 Corinthians 12:1, 4"),
    ('v', "[1] Now concerning spiritual gifts, brethren, I would not have you ignorant."),
    ('v', "[4] Now there are diversities of gifts, but the same Spirit."),
    ('p', "Now concerning spiritual gifts, Paul's speaking: things of the Spirit, brethren, I would not have you ignorant. So because that word there is pneumatikos, right, remember that? Pneuma, pneumatikos, that is, the things of the Spirit, what is given by the Spirit. So Paul is now to explain those things to us."),
    ('p', "Now you shouldn't learn things primarily by experience; experiences will differ. As you see in the Book of Acts, things happened differently. But if you learn by doctrine, you'll be able to guide your experiences and explain them as well."),
    ('p', "So Paul calls them pneumatikos, and in verse 4, okay, so what Paul has done for us, Paul has taken the veil away from the Book of Acts so we can see what happened literally. That in verse 4, look at verse 4. Look at verse 4: 'Now,' come on let's take verse 4 together, let's go: 'there are diversities of gifts, but the same Spirit.'"),
    ('p', "What is the word gifts that Paul used there? Charisma, which is what? Charisma means an endowment, an ability, right? Charisma means an endowment, an ability. It's used for people, you know when you say that guy is gifted, right? He's gifted in speaking. When you see a football player like Lionel Messi, you know he is gifted. When you see Cristiano Ronaldo, he is gifted. When you see other players, they are players. No but I'm not joking."),
    ('p', "You know there are some teams you play with and if Lionel Messi is in the team and they have just five players fit, they can still beat the other team. Just give him the ball often. You know if you're a footballer, if you've played football before, you take the ball from the defense, if you're like the distributor of the team, what you call the playmaker, you look around your team, who can I give the ball at this particular point in time? You see nobody. You look away. You see, I should stay in faith. So let me just, I should stay in faith, right?"),
    ('p', "But you see Lionel Messi or you see Cristiano Ronaldo, you just pass the ball to him, you don't care what's going to happen, because you know he's gifted! But there are some folks, you pass to them, you say move to the left, move to the right, because you're not even sure he knows what to do with it!"),

    # Charisma as Natural Endowment
    ('p', "So charisma is a natural endowment. It was used, it was used by the Greeks. In fact, Plato and Aristotle, who were major artists of words before Jesus came, they used it for natural talents. So charisma is what becomes natural. So we can say there are diversities of what is natural, but the same Spirit."),
    ('p', "Don't forget, the Spirit of God has become the nature of the believer, can you remember that? So these are diversities. And he mentions three. You know, Howard Carter or Howard Carter majorly said to us that there are three groupings of the gifts of the Spirit, and I agree with him: they are the utterance gifts that say something, and then the revelation gifts that show something, and you have the power gifts that do something. Remember that? Remember that?"),
    ('p', "But the keyword, all right, the keyword is that these things are abilities. So the things we read in the Book of Acts were not experiences; they were abilities that produced experiences."),
    ('p', "As you are right now, you are naturally endowed by the Spirit of God to do supernatural things, naturally! Now when I say naturally, I'm not referring to your human makeup, I'm referring to your spiritual makeup. Is that very clear? Come on guys, is that very clear? Good."),

    # Gifts of Healings - Diversities of Endowments
    ('p', "Now let's just quickly talk about the gifts of healings quickly, because I hope I can rush through that. In verse 9, he says:"),
    ('ref', "1 Corinthians 12:9, 28, 30"),
    ('v', "[9] To another faith by the same Spirit; to another the gifts of healing by the same Spirit;"),
    ('v', "[28] And God hath set some in the church, first apostles, secondarily prophets, thirdly teachers, after that miracles, then gifts of healings, helps, governments, diversities of tongues."),
    ('v', "[30] Have all the gifts of healing? do all speak with tongues? do all interpret?"),
    ('p', "'To another faith by the same Spirit; to another the gifts of healing by the same Spirit.'"),
    ('p', "Now he uses the word gift, don't forget he's discussing gifts, he used the word gifts of healing three times: verse 9, that's the only gift he repeats the word gifts again, verse 9; and then verse 28, he said he set some in the church; then verse 30, 'have all the gifts of healing?'"),
    ('p', "So it's a fact that the ability to heal, which is in Christ, is in diversities. Now listen: so that means in the endowment for healing are diverse endowments. I'll explain to you what I mean."),
    ('p', "Now I like art. I like good dramas, very good ones, that don't necessarily have to have 'Babalawo' in them and all that, you know, but very good dramas that you know, and I really love to watch them. I remember one I watched recently, I loved the drama, October 1, quite good, because that guy, what's his name again, Kunle Afolayan, is quite one of the best produced in Africa, quite solid."),
    ('p', "Now in watching drama, I have this aversion for a typecast. A typecast is somebody that if you see him as a bad guy, if he's doing good guy you think he's a bad good guy! Just a typecast, just one way. There are some folks, they can only act rich man."),
    ('p', "When I used to do dramas, I almost became a typecast because I always acted the guy that will miss the rapture! Or a backsliding Christian, something like that, something terrible like that, backsliding Christian, miss the rapture, because they felt I could do it very well. You know, I could miss the rapture very easily! You want to pray for me? They used to pray for me."),
    ('p', "Okay so, but there are some artistes, you know if you've seen what do you call it, Terminator, you have Arnold Schwarzenegger, Terminator 1 he was a bad guy, Terminator 2 he became the good guy. So he can act this way, he can act that way. That means that kind of actor is endowed to play different parts, though he's still acting, you get it?"),
    ('p', "Now the gifts of healing mean that the believer is endowed to minister healing differently. Sometimes you can get used to just one way. He's endowed to minister healing differently. Are you listening? Come on guys, are you listening? Okay."),

    # Laying on of Hands & Matthew 8
    ('p', "Now the first one we want to examine today, and probably a bit next week, is the laying on of hands. I've heard people teach how to receive healing, they say you receive healing by faith, by the word, there's laying on of hands and gifts of healing. It's ignorance, with due respect to those who teach that."),
    ('p', "Now gifts of healing already shows you, you know we all said that it can't be gifts of healing based on the kind of sickness, can you remember that? Can you remember that? Good. We said it's gifts of healing because the word gifts already have to do with ministry, so it has to be the ministries to the sick in different ways. Do you understand that? Do you understand that? Okay."),
    ('p', "So now, the first one we want to examine is laying on of hands. Laying on of hands is a gifts of healing. Mark 16: 'They shall lay hands on the sick, and the sick shall recover.' They, that's the sick, I always like to say and the sick, because the problem there is the word they: they shall lay hands and they shall recover, you know, so I always like to say and the sick shall recover. I hope you understand that? Okay, thank you."),
    ('p', "Now what does it mean to lay hands? Lay hands on the sick and the sick shall recover, what does it mean? Look at Matthew 8. Matthew 8. Here's Jesus. Matthew 8."),
    ('p', "In the Book of Matthew, the healing ministry of Jesus was mentioned, pay attention now, his healing ministry was mentioned in Matthew 4 verse 23 and 24, but how it was wasn't told. He said they brought unto him verse 24 all sick people that were taken with divers diseases, torments, and those that were possessed, and they were lunatics, and the Bible says they were healed. He didn't tell us how. But in Matthew 8, we now find specifics."),
    ('p', "He said and there was a man, okay, who was a leper, said 'Lord, if thou wilt, thou canst make me clean.' And Jesus put forth what? Uh? Verse 3, Matthew 8, he put forth what? And did what? Touched him. Notice that: he put forth his hand and touched, and said, 'I will; be thou clean.'"),
    ('ref', "Matthew 8:2-3"),
    ('v', "[2] And, behold, there came a leper and worshipped him, saying, Lord, if thou wilt, thou canst make me clean."),
    ('v', "[3] And Jesus put forth his hand, and touched him, saying, I will; be thou clean. And immediately his leprosy was cleansed."),
    ('p', "Now go somewhere further. In verse 14:"),
    ('ref', "Matthew 8:14-15"),
    ('v', "[14] And when Jesus was come into Peter's house, he saw his wife's mother laid, and sick of a fever."),
    ('v', "[15] And he touched her hand, and the fever left her: and she arose, and ministered unto them."),
    ('p', "So the very first Pope had a wife! Is that not true? Very first Pope, he had a wife. 'And he did what? Uh?' Now pay attention: he touched what? Her hand. 'And the fever left her: and she arose, and ministered unto them.'"),
    ('p', "He touched her hand. You know the first one, we don't know what he touched, but the second time he touched her hand. Now many of us, I'm going to show you in a moment, many of us think laying on of hands is on the head. So even if someone says he has back pain, you put your hand on the head! Are you listening? Put your hand on the head. He said he touched her hand. Praise the Lord! Are you there?"),
    ('p', "He put his hand, just touched her hand, that's what he did. Now did he lay his hands on them? Did he lay his hands on her, uh, by what? Praise the Lord!"),

    # 1 Timothy 5:22 vs Ministry Ordination
    ('p', "Now go to 1 Timothy 5. 'Lay hands,' verse 22:"),
    ('ref', "1 Timothy 5:22"),
    ('v', "[22] Lay hands suddenly on no man, neither be partaker of other men's sins: keep thyself pure."),
    ('p', "'Lay hands suddenly on no man.'"),
    ('p', "Now I've heard people say that word lay hands on no man means when you're in a meeting you can't just lay hands on people anyhow. Rubbish! Where it says lay hands suddenly on no man here, neither be a partaker of other men's sins, keep thyself pure: this was a verse they used for me when I laid hands on the sick in 1992, that I had polluted myself with people's sins, because I didn't know who they were."),
    ('p', "It's good to know your Bible well enough. Paul's discussion here was about ordination. The word lay hands suddenly on no man is the word hastily, suddenly, hastily. It's the word tacheos in the Greek, t-a-c-h-e-o-s, it means to do it hastily. That is, you, for instance he said to him, do not ordain anyone to ministry except the person is blameless, husband of one wife, okay, is able to teach. So say lay hands suddenly on no man."),
    ('p', "Look at Acts 13, Acts 13: Paul and Barnabas. Verse 3:"),
    ('ref', "Acts 13:3"),
    ('v', "[3] And when they had fasted and prayed, and laid their hands on them, they sent them away."),
    ('p', "They laid their hands on them for ministry. Okay? Are you there? Are you following? Good."),
    ('p', "So now, so the laying on of hands for ministry is not touch; it's one that involves to present people for ministry, which will involve you know what you call ordination, appointment, or consecration to the work of ministry."),

    # Mark 8:22-25 - Touching the Sick Deliberately
    ('p', "Now in laying hands on the sick, you'll discover that what Jesus did was to touch the sick. Mark 8:22. Are you listening? Are you listening? Okay, 22. Sorry, Mark 8, sorry for that, Mark 8:22:"),
    ('ref', "Mark 8:22-25"),
    ('v', "[22] And he cometh to Bethsaida; and they bring a blind man unto him, and besought him to touch him."),
    ('v', "[23] And he took the blind man by the hand, and led him out of the town; and when he had spit on his eyes, and put his hands upon him, he asked him if he saw ought."),
    ('v', "[24] And he looked up, and said, I see men as trees, walking."),
    ('v', "[25] After that he put his hands again upon his eyes, and made him look up: and he was restored, and saw every man clearly."),
    ('p', "'And he cometh to Bethsaida; and they bring a blind man unto him, and besought him to touch him.' That's what, to touch him."),
    ('p', "'And he took the blind man by the hand, and led him out of the town; and when he had spit on his eyes, and put his hands upon him, he asked him if he saw ought. And he looked up, and said, I see men as trees, walking.' That means the man was getting healed gradually."),
    ('p', "'After that he put his hands again where? Where? Upon his eyes, and made him look up: and he was restored, and saw every man clearly.'"),
    ('p', "Okay, so in touching the sick, which in more context we're going to see that, it's advisable, all right, to put your hands where, where? Where the pain is. Now I'm going to go to this later."),
    ('p', "Now if you have a woman, and you're a man, and for instance you have some women that will tell you their pain is somewhere that is private, you don't say pastor said, uh-uh, no, no, brother, no! All right? There's a way to do it that I'm going to explain that to you."),
    ('p', "But you see, he laid hands where there was the pain. He put his hands there, where there was the pain. So laying on of hands is not laying hands on the head, are you getting this? Are you following it? Now come back again: Mark 10."),

    # Mark 10 - Touching Young Children
    ('p', "Mark 10. Verse 13:"),
    ('ref', "Mark 10:13, 16"),
    ('v', "[13] And they brought young children to him, that he should touch them: and his disciples rebuked those that brought them."),
    ('v', "[16] And he took them up in his arms, put his hands upon them, and blessed them."),
    ('p', "'And they brought young children to him, that he should touch them: and his disciples rebuked those that brought them.' Okay? And look at verse 16: 'And he took them up in his arms, put his hands upon them, and blessed them.' So to touch."),
    ('p', "So that means it's a deliberate, so the word lay in the Greek is a deliberate touch."),
    ('p', "Now you know there's a way, let me say someone so, Maiwada, where he's sometimes you want to leave the auditorium of the church, and he's moving before me, you are too efficient, can you calm down? And he's moving, and I'm like, excuse me. You know, I've touched him, but I'm not laying hands on him, but I touched him."),
    ('p', "And I can also touch him like this: what is it now, you know, right? Am I not touching him? I can also touch him like this: go sit down! All right, so I can lay hands like that. You can you can also touch someone like this: how are you, and then you hug yourself, say how are you doing? So you can't say you're laying hands."),
    ('p', "So the word lay hands is a deliberate touch. Now what is that deliberate touch for? It is to minister healing. It's a deliberate touch. You'll find out the word lay hands, you find that they laid hands upon Paul, upon Peter, the soldiers, they laid hands. It's also deliberate, you know I said deliberate touch. Not for turkey, they laid hands for what? To arrest! That's a deliberate touch."),
    ('p', "So laying on of hands is a deliberate touch. While our own is to minister, we want to arrest the sickness! To minister to the sick. So it is not to put hands upon the head of people, do you understand it? Do you understand? It's not to put hands upon the head of people. It is a deliberate touch on the sick, which many times is on the part where there is the pain."),

    # Mark 6: Deliberate Touch in Ministry
    ('p', "Look at Mark 6. So laying on of hands is a deliberate touch. So if someone has a leg pain, where is it advisable I lay hands? On the leg. Praise the Lord! Amen. We're going to go into the specifics next week. But you see, it's a deliberate touch, that is, I am touching the person with something."),
    ('p', "You see what I said? If Maiwada can be ill, praise God he's healed, but someone is coming to the service and I say how you doing, and we shake hands, and he goes. Then he comes back and says sir, I want you to pray for me, I'm sick in my body. I touched him earlier, and that was just how you doing. Now what do I do? Uh? I now deliberately, so that laying on of hands is to lay hands with the power of God. It's deliberate!"),
    ('p', "Now you know for instance, I don't think that when Peter's mother-in-law got up, she just looked at Jesus from afar. She must have you know moved close to him, served him food, and probably had said you're blessed."),
    ('p', "And I don't think that when he touched her the second time she started shaking, ha brother master let me heal you, what's going on the power is here again, you know, no! Because that's not laying, even though he must have touched her, or shook her hand, but that's not, that's not a deliberate touch, do you understand that? Do you understand that? So laying on of hands is deliberate. Praise the Lord!"),

    # Mark 6 & Acts 28 - Healing in Unbelief and Confinement
    ('p', "Now see this: Mark 6. The effect, the effectiveness of laying on of hands can be seen here. In Jesus's own town, they actually dishonored him. The Bible says they were offended at him, in verse 3, Mark 6. And he said, Jesus said:"),
    ('ref', "Mark 6:4-5"),
    ('v', "[4] But Jesus said unto them, A prophet is not without honour, but in his own country, and among his own kin, and in his own house."),
    ('v', "[5] And he could there do no mighty work, save that he laid his hands upon a few sick folk, and healed them."),
    ('p', "'A prophet is not without honour, but in his own country, and among his own kin, and in his own house. And he could there do no mighty work,' notice this, 'he could do no mighty work, save that he laid his hands upon a few sick folk, and healed them.'"),
    ('p', "So even in the midst of unbelief, laying on of hands can work! It still laid hands."),
    ('p', "So in Matthew 20, sorry Acts 28 quickly, we're going to explain this a little bit more. Acts 28. This is Paul in captivity. Notice, he's in captivity, he's under imprisonment in a sort, because he's under arrest, he's being taken from place to place, yield to appear before Caesar. Now notice this: Acts 28 verse 7:"),
    ('ref', "Acts 28:7-9"),
    ('v', "[7] In the same quarters were possessions of the chief man of the island, whose name was Publius; who received us, and lodged us three days courteously."),
    ('v', "[8] And it came to pass, that the father of Publius lay sick of a fever and of a bloody flux: to whom Paul entered in, and prayed, and laid his hands on him, and healed him."),
    ('v', "[9] So when this was done, others also, which had diseases in the island, came, and were healed."),
    ('p', "'In the same quarters were possessions of the chief man of the island, whose name was Publius; who received us, and lodged us three days courteously. And it came to pass, that the father of Publius lay sick of a fever and of a bloody flux: to whom Paul entered in, and prayed, and laid his hands on him, and healed him.'"),
    ('p', "Now notice something, pay attention: 'So when this was done,' what was done? Okay, good. 'Others also, which had diseases in the island, came, and were healed.'"),
    ('p', "Did you observe a similarity in the Matthew 8 account and this account? Like laying on of hands is the way to start. When Jesus laid hands in Matthew 8, verse 16 and 17, that's where many people now came and they were healed. The same thing here, I'm going to show you the sequence next week."),
    ('p', "So laying on of hands is a primary way to start ministering to the sick. I mean both of them: both Jesus and this man, Jesus was in a place he wasn't honored, Paul was in confinement, and he laid hands on the sick."),

    # The Endowment: Power in Action
    ('p', "So what is the endowment here? What is it? You know it says gifts of healing, right? So what's the endowment? Uh? Your hands? What do you call it? Ability to what? To put your hands on the sick? Everybody can do that! Doctors do it every time, giving injection, right? Say what's wrong with you, say what's wrong with you. Are they not laying hands?"),
    ('p', "So what is the, what is the gift? To minister the power of God. What exactly is the gift? Mantle? I didn't hear what he's saying. Uh? Coming through the body? I don't understand that. I got you, I was just kidding."),
    ('p', "Now it's a, laying on of hands will be what: a deliberate touch. Can you use your leg? And that is your hand. I trapped you! Praise the Lord! Okay, we'll get to that anyway. So what is the gift there? Okay, go and think about it. I won't say it. I've said it already, I won't repeat it. So you can touch people deliberately, okay, to minister to the sick, to minister divine healing unto them."),

    # Acts 3 & Acts 8: Action and Diverse Administrations
    ('p', "Now notice something, and we're close here now. Notice something, and I want to show you quickly: in Acts 3, in the case of Peter, notice how what they do is described. Acts 3, that man at the Beautiful Gate, Bible says:"),
    ('ref', "Acts 3:6-8"),
    ('v', "[6] Then Peter said, Silver and gold have I none; but such as I have give I thee: In the name of Jesus Christ of Nazareth rise up and walk."),
    ('v', "[7] And he took him by the right hand, and lifted him up: and immediately his feet and ankle bones received strength."),
    ('v', "[8] And he leaping up stood, and walked, and entered with them into the temple, walking, and leaping, and praising God."),
    ('p', "'And he took him by the right hand, and lifted him up: and immediately his feet and ankle bones received strength.'"),
    ('p', "Where did the strength come from? It came through Peter's what? Uh? Through Peter's what? Touch, right?"),
    ('p', "Now look at Acts 8 verse 5:"),
    ('ref', "Acts 8:5-7"),
    ('v', "[5] Then Philip went down to the city of Samaria, and preached Christ unto them."),
    ('v', "[6] And the people with one accord gave heed unto those things which Philip spake, hearing and seeing the miracles which he did."),
    ('v', "[7] For unclean spirits, crying with loud voice, came out of many that were possessed with them: and many taken with palsies, and that were lame, were healed."),
    ('p', "'Then Philip went down to the city of Samaria, and preached Christ unto them. And the people with one accord gave heed unto those things which Philip spake, hearing and seeing the miracles which he did.'"),
    ('p', "So that means the healing or the power is in your action. It's in your action! Lay hands on the sick. It didn't say go and declare the sick healed, say lay hands on the sick. Same thing like Paul: Paul healed them, that's what the Bible says, Acts 28. So it is in the action."),
    ('p', "So that means the diversity is the action, diverse actions. Praise the Lord! Praise the Lord! Do you understand that? It's in the action."),

    # Conclusion & Prayer
    ('p', "So in our studying the gifts of healing, we're going to study the action, we're not done with laying on of hands, we're going to continue next week. We're going to study the action and see how it's done."),
    ('p', "Notice in the instances, those folks received, they received from what? The touch. It's from the touch that they received. Praise the Lord!"),
    ('p', "You're blessed? Let's stand to our feet. Let's pray in the Holy Ghost, sing in the Holy Ghost, sorry. Sing in the Holy Ghost."),
    ('p', "[Singing in the Spirit / Tongues]")
]

def build_docx(output_path):
    doc = Document()
    
    # 1.0 inch margins
    section = doc.sections[0]
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    
    # Header / Title matching repository format
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    t_run = title_p.add_run("In the Spirit by the Spirit – Power Gifts - Track 2")
    t_run.font.name = "Lexend"
    t_run.font.size = Pt(14)
    t_run.font.bold = True
    
    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_before = Pt(0)
    sub_p.paragraph_format.space_after = Pt(14)
    s_run = sub_p.add_run("Teaching Transcript")
    s_run.font.name = "Lexend"
    s_run.font.size = Pt(10)
    s_run.font.italic = True
    s_run.font.color.rgb = RGBColor(100, 100, 100)
    
    for elem_type, text in ELEMENTS:
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.font.name = "Lexend"
        
        if elem_type == 'ref':
            # Scripture reference heading: Bold
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(2)
            run.font.size = Pt(10)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 0, 0)
        elif elem_type == 'v':
            # Scripture verse line: Bold
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(2)
            run.font.size = Pt(9.5)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 0, 0)
        else:
            # Spoken discourse: Regular, natural spacing
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(5)
            p.paragraph_format.line_spacing = 1.15
            run.font.size = Pt(10)
            run.font.bold = False
            run.font.color.rgb = RGBColor(30, 30, 30)
            
    doc.save(output_path)
    print(f"Docx built: {output_path}")

def build_txt(output_path):
    lines = []
    lines.append("In the Spirit by the Spirit – Power Gifts - Track 2")
    lines.append("Teaching Transcript")
    lines.append("=" * 60)
    lines.append("")
    
    for elem_type, text in ELEMENTS:
        if elem_type == 'ref':
            lines.append(f"**{text}**")
        elif elem_type == 'v':
            lines.append(f"**{text}**")
        else:
            lines.append(text)
            lines.append("")
            
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(lines))
    print(f"Txt built: {output_path}")

if __name__ == "__main__":
    out_docx_raw = os.path.join("In the Spirit by the Spirit Transcript", "Raw", "In the Spirit by the Spirit – Power Gifts - Track 2_Raw.docx")
    out_docx_raw_alt = os.path.join("In the Spirit by the Spirit Transcript", "Raw", "In the Spirit by the Spirit – Power Gifts - Track 2 (Structured Verbatim).docx")
    out_docx_root = "In the Spirit by the Spirit – Power Gifts - Track 2_Raw.docx"
    out_txt_raw = os.path.join("In the Spirit by the Spirit Transcript", "Raw", "In the Spirit by the Spirit – Power Gifts - Track 2_Raw.txt")
    
    # Save root first
    build_docx(out_docx_root)
    
    # Try raw docx; if locked by Word, save to alt
    try:
        build_docx(out_docx_raw)
    except PermissionError:
        print(f"Warning: '{out_docx_raw}' is currently open in Word. Saving to alternate path '{out_docx_raw_alt}'")
        build_docx(out_docx_raw_alt)
        
    build_txt(out_txt_raw)
    print("Master build complete!")
