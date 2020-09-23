<div align="center">

## Building a SQL Search Query\-\-Easily


</div>

### Description

This article explains an alternative to the boring string parse/loop method of building a search query for your website.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Primoris Software](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/primoris-software.md)
**Level**          |Intermediate
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/primoris-software-building-a-sql-search-query-easily__4-6701/archive/master.zip)





### Source Code

```
<div class=Section1>
<h1 align=center style='text-align:center'><span style='font-size:18.0pt;
mso-bidi-font-size:16.0pt'>Building a Better Query<o:p></o:p></span></h1>
<h2><span style='font-size:12.0pt;mso-bidi-font-size:14.0pt'>What kind of Query
does this apply to?<o:p></o:p></span></h2>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:10.0pt'>This
tutorial is meant to cover the type of queries where someone is running a
single or multiple keyword search on your web site (or, I suspect this could be
used in software development as well, though I haven't used it yet).<span
style="mso-spacerun: yes">  </span><o:p></o:p></span></p>
<p class=MsoBodyText>For example, you have a knowledge base, and a customer is
looking up information on a printing bug.<span style="mso-spacerun: yes"> 
</span>So they go to your &quot;Search&quot; field and type &quot;print bug
epson&quot;.<span style="mso-spacerun: yes">  </span>Sound easy?<span
style="mso-spacerun: yes">  </span>You're not saying, 'Oh, no problem, the query
would look like SELECT * FROM TABLE WHERE field LIKE '%&quot; &amp; textbox
&amp; &quot;%'&quot;', are you?<span style="mso-spacerun: yes">  </span>The
problem is, people many times type their keywords in an order different than
that found in your knowledge base, and certainly there's the possibility of
other words between the keywords.</p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:10.0pt'>So
how do you overcome this?<span style="mso-spacerun: yes">  </span>One solution
(widely used) is to use a loop.<span style="mso-spacerun: yes">  </span>Parse
all of the words with a commonly used parsing function.<span
style="mso-spacerun: yes">  </span>Sure, you can download one from one of the
free code places on the web, but then you have to find it and figure out how to
use it.<span style="mso-spacerun: yes">  </span>You can write it yourself, but
I promise you it is one of the most boring pieces of code you will ever write.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:10.0pt'>After
all of the keywords are parsed into an array, build your query by looping
through all of them.<span style="mso-spacerun: yes">  </span>Fast?<span
style="mso-spacerun: yes">  </span>Easy?<span style="mso-spacerun: yes"> 
</span>Efficient?<span style="mso-spacerun: yes">  </span>Kind of.<span
style="mso-spacerun: yes">  </span>But here's a better, more organized way.<o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<h2><span style='font-size:11.0pt;mso-bidi-font-size:14.0pt'>The Alternative
Way to Build the Query<o:p></o:p></span></h2>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:10.0pt'>The
alternative comes in two parts.<span style="mso-spacerun: yes">  </span>Step
one is to remove all extraneous white space from the string.<span
style="mso-spacerun: yes">  </span>It <i style='mso-bidi-font-style:normal'>probably</i>
doesn't have any, but you never know.<span style="mso-spacerun: yes"> 
</span>So, do this on the client side with a function that loops something
like,<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:10.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:1'>          </span>while
(strObj.search(&quot;<span style="mso-spacerun: yes">  </span>&quot;) != -1)<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:1'>          </span>{<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:2'>                   </span>strObj
= strObj.replace(&quot;<span style="mso-spacerun: yes">  </span>&quot;,&quot;
&quot;);<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:1'>          </span>}<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:1'>          </span>return
strObj.toString()<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:8.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></code></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:10.0pt'>In
case you didn't catch it, in the while condition there is a string with two
spaces, and in the replace function, the string with two spaces is replaced
with one.<span style="mso-spacerun: yes">  </span>Until there is only one space
between (or in front or behind of) every keyword, keep stripping it down.<span
style="mso-spacerun: yes">  </span><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:10.0pt'>Now
you have a nicely organized string of keywords with one space max between,
before, or after the words.<span style="mso-spacerun: yes">  </span>On the
search page, you may also have conditions like &quot;case sensitive&quot;,
matches per page, etc, but we won't go over that here (if you want me to cover
that at some point, simply send me a note and I'll try to do it).<o:p></o:p></span></p>
<p class=MsoBodyText>Pass this string to your next active server page.<span
style="mso-spacerun: yes">  </span>For me, did you notice that I returned the
string from a function in the code above?<span style="mso-spacerun: yes"> 
</span>I stuck that return value back in the original text box.<span
style="mso-spacerun: yes">  </span>It's up to you to decide how you want to
pass the string back.</p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:10.0pt'>Either
way, now the ASP has to build the search string.<span style="mso-spacerun:
yes">  </span>Here's the good part—step two.<span style="mso-spacerun: yes"> 
</span>Now that your string has only single spaces, you can strip the first and
last spaces, and run a replace on all of the remaining.<span
style="mso-spacerun: yes">  </span>Let's build this step by step (the way I do
on my page).<o:p></o:p></span></p>
<p class=MsoNormal><code><span style='font-size:8.0pt;mso-bidi-font-size:10.0pt;
font-family:"Courier New";mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'>Function buildSQL( strText )<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:1'>          </span>Dim
selectClause<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:1'>          </span>Dim
fromClause<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:1'>          </span>Dim
whereClause<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:2'>                   </span><o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:1'>          </span>selectClause
= &quot;SELECT<span style="mso-spacerun: yes">  </span>* &quot;<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:1'>          </span>fromClause
= &quot;FROM table &quot;<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:1'>          </span>whereClause
= &quot;WHERE &quot;<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:1'>          </span>Select
Case Request.QueryString(&quot;optAll&quot;)<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:2'>                   </span>Case
1 ' Match ALL keywords<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:2'>                   </span><span
style="mso-spacerun: yes">    </span>whereClause = whereClause &amp; &quot;
(field LIKE '%&quot; &amp; _ <o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'>Replace( Trim( strText ), &quot;
&quot;, &quot;%' AND field LIKE '%&quot;) &amp; &quot;%')&quot;<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:2'>                   </span><span
style="mso-spacerun: yes">  </span><span style='mso-tab-count:1'>        </span>Case
2 ' Match ANY keywords<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:2'>                   </span>whereClause
= whereClause &amp; &quot; (field LIKE '%&quot; &amp; _ <o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'>Replace( Trim( strText ), &quot;
&quot;, &quot;%' OR field LIKE '%&quot;) &amp; &quot;%')&quot;<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:1'>          </span>End
Select<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:1'>          </span>'Response.Write(
selectClause &amp; fromClause &amp; whereClause )<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'><span style='mso-tab-count:1'>          </span>buildSQL
= selectClause &amp; fromClause &amp; whereClause<o:p></o:p></span></code></p>
<p class=MsoNormal style='mso-pagination:widow-orphan lines-together'><code><span
style='font-size:7.0pt;mso-bidi-font-size:10.0pt;font-family:"Courier New";
mso-bidi-font-family:"Times New Roman"'>End Function<o:p></o:p></span></code></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:10.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><i style='mso-bidi-font-style:normal'><span
style='font-size:8.0pt;mso-bidi-font-size:10.0pt'>So</span></i><span
style='font-size:8.0pt;mso-bidi-font-size:10.0pt'> you're replacing all of the
middle spaces with a SQL 'and' statement.<span style="mso-spacerun: yes"> 
</span>In plain English, if your search phrase is &quot;print bug&quot;, this
now becomes &quot;'%print%' AND field LIKE '%bug%'&quot; when you concatenate
the leading and trailing %'s and quotes (this is for Microsoft Access drivers,
other drivers may use different wildcards)--so just append this phrase to the
&quot;WHERE field LIKE &quot; phrase, and you're in business.<span
style="mso-spacerun: yes">  </span>I've built gigantic search phrases with this
method before with little coding, and little server load.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:10.0pt'>Voila!<span
style="mso-spacerun: yes">  </span>An instant search query!<span
style="mso-spacerun: yes">  </span>No tiresome string parsing or looping.<span
style="mso-spacerun: yes">  </span>One final question you may have is,
&quot;what if the user separates the keywords with commas or hyphens or...&quot;.<span
style="mso-spacerun: yes">  </span>No problem!<span style="mso-spacerun: yes"> 
</span>Just put client-side code in to convert all hyphens, commas, etc. to
white space.<span style="mso-spacerun: yes">  </span>Put this <i
style='mso-bidi-font-style:normal'>before</i> the function that strips the
white-space down to one.<span style="mso-spacerun: yes">  </span>String:
Normalized.<span style="mso-spacerun: yes">  </span>So that's how it's
done.<span style="mso-spacerun: yes">  </span>If you have any questions, I
would be happy to explain further--just send me a note on Planet Source Code.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:10.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
</div>
```

