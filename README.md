![Hotpho Icon](https://github.com/nguyenkvvn/Hotpho-VSTO/blob/master/Resources/Artboard%201small.png)
## Hotpho-VSTO
An experimental and educational proof-of-concept VSTO Word plugin of demonstrate defeating online content originality checkers. Will not survive a through inspection by design.

## Installation

(No executable binary file will be provided to provide a barrier of entry; intended to deter unwarranted misuse and widespread abuse of the concept.)

1. Clone the Hotpho repository.
2. Open the Hotpho.sln in Visual Studio.
3. Build the plugin.
4. Word will launch with the plugin installed as a debug build.

## Usage

1. Install the plugin.
2. Open Word 2013.
3. Highlight the range of text to obfuscate.
4. Click on the "ADD-INS" tab on the Ribbon in Word.
5. With the range of text highlighted, click the "Obfuscate Selection" button in the Hotpho plugin's ribbon section.

## Contributing

*(Kept as is from template.)*
1. Fork it!
2. Create your feature branch: `git checkout -b my-new-feature`
3. Commit your changes: `git commit -am 'Add some feature'`
4. Push to the branch: `git push origin my-new-feature`
5. Submit a pull request :D

## History

The application's purpose for me was to explore the venue of developing plugins for Microsoft Office using the VSTO platform. The platform is relatively niche, and was a pleasant learning experience. The idea sprung about was when I noticed while submitting a document on a certain classroom platform at the local college that it had a "originality check" score next to my file submission. I didn't give a hoot what the score was (since I cite my sources and enjoy writing enough to not plagiarize), but now I was curious as to what made this peculiar feature tick. Research revealed that similar services and platforms highlight crossing sentence structure and content to the submitted document. One source in particular noted how students would add "filler" content in between page breaks and words- none said anything about characters though (which would be laborious).

## Defending against this plugin

There are a number of ways to defeat this plugin. They all either involve considerable effort in terms of processing volume, labor, or some clever implementations of solutions:
1. The students actually write the paper themselves and cite the sources. (Pros: Honesty and guaranteed originality check pass. Cons: requires effort.)
2. The teachers carefully read the paper students submit. (Pros: Careful and accurate detection of linguistic discrepancies among their student. Cons: requires effort. Not feasible when delegated to TAs or for larger college classrooms.)
3. Teachers check for the originality flag. Consider the false positive paradox- a paper, sufficiently lengthy enough, with a 100% originality score should be worth a read, and at the minimum, an inspection, as the paper could be a wholesome and thoughtful read or loaded with complete gibberish and rubbish to fool the checker. (Realistically, completely original papers should not have 100% originality score as they must pull their linguistic and syntaxical influences from *somewhere*. (Pros: Filtered screening, talent discovery. Cons: requires effort and checking.)
4. Revealing the "Hidden" text property in word, or showing raw text outside of Word (aka a service). (Pros: Guaranteed screening and proof of dishonesty. Cons: not feasible for bulk submissions.)
5. Teachers and graders copy/paste from the document into an online checker. (Pros: copy/pasting text with hidden content only copies visible text, effectively defeating any hidden text obfuscation. Cons: not feasible in bulk submissions)
6. Configure the originality checker to ignore all hidden content. (Pros: logically sound option to use. Cons: third-party feature submission & development required)

The Hotpho plugin got its name as a reference to a track on certain underrated band's album.

## Disclaimer

Hotpho and its developer is not responsible or liable for any injuries, penalties, damages, or transgressions you commit or incur before, after, or during your use of this application. You are solely responsible for your usage and the outcome of your actions for the use of this plugin. **The developer, nor any parties involved in development, condone or approve of using this tool to cheat, deceive, or mislead any persons or entities.** Dishonesty is, quite frankly, frowned upon by the majority of the academic and professional community, as well as the developers themselves; it would be reprehensible to use this tool maliciously. Use this plugin at your own risk. YOU HAVE BEEN WARNED. (Read the license below.)

## License

>MIT License
>
>Copyright (c) 2017 Vinh Nguyen
>
>Permission is hereby granted, free of charge, to any person obtaining a copy
>of this software and associated documentation files (the "Software"), to deal
>in the Software without restriction, including without limitation the rights
>to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
>copies of the Software, and to permit persons to whom the Software is
>furnished to do so, subject to the following conditions:
>
>The above copyright notice and this permission notice shall be included in all
>copies or substantial portions of the Software.
>
>THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
>IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
>FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
>AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
>LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
>OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
>SOFTWARE.
