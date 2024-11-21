# quizXL
quizXL is a macro-enabled Microsoft Excel application that can generate and export Classic Quizzes (QTI version 1.2 format ZIP-file), and both learning outcomes (.CSV), to the CANVAS LMS digital learning platform

## Features
- Full support for all 9 question types in CANVAS Classic Quizzes that have automatic grading.
  - **true_false_question**: called '[True/False](https://community.canvaslms.com/t5/Instructor-Guide/How-do-I-create-a-True-False-quiz-question/ta-p/927)' in CANVAS
  - **short_answer_question**: called '[Fill in the blank](https://community.canvaslms.com/t5/Instructor-Guide/How-do-I-create-a-Fill-in-the-Blank-quiz-question/ta-p/889)' in CANVAS
  - **multiple_choice_question**: called '[Multiple choice](https://community.canvaslms.com/t5/Instructor-Guide/How-do-I-create-a-Multiple-Choice-quiz-question/ta-p/682)' in CANVAS
  - **multiple_answers_question**: called '[Multiple answers](https://community.canvaslms.com/t5/Instructor-Guide/How-do-I-create-a-Multiple-Answers-quiz-question/ta-p/924)' in CANVAS
  - **fill_in_multiple_blanks_question**: called '[Fill in multiple blanks](https://community.canvaslms.com/t5/Instructor-Guide/How-do-I-create-a-Fill-in-Multiple-Blanks-quiz-question/ta-p/923)' in CANVAS
  - **multiple_dropdowns_question**: called '[Multiple dropdowns](https://community.canvaslms.com/t5/Instructor-Guide/How-do-I-create-a-Multiple-Dropdown-quiz-question/ta-p/916)' in CANVAS
  - **matching_question**: called '[Matching](https://community.canvaslms.com/t5/Instructor-Guide/How-do-I-create-a-Matching-quiz-question/ta-p/918)' in CANVAS
  - **numerical_question**: called '[Numerical answer](https://community.canvaslms.com/t5/Instructor-Guide/How-do-I-create-a-Numerical-Answer-quiz-question/ta-p/919)' in CANVAS
  - **calculated_question**: called '[Formula question](https://community.canvaslms.com/t5/Instructor-Guide/How-do-I-create-a-Formula-quiz-question-with-a-single-variable/ta-p/920)' in CANVAS
- Enables you to rapildy build a large pool of quiz questions. Each worksheet is a different quiz.
- Export to CANVAS by pressing CTRL+E (Export). The macro checks your input for errors, then exports a QTI-format ZIP file (QTI version 1.2 for Classic Quizzes). [Here are instruction on how to import the QTI-format ZIP-file into CANVAS](https://community.canvaslms.com/t5/Instructor-Guide/How-do-I-import-quizzes-from-QTI-packages/ta-p/1046). It gets imported as an unpublished Classic Quiz.
- Supports quiz groups. Each group can cover specific learning goals. You can therefore build randomized quizzes with large pool of questions in each group. This ensures that each repetition of the quiz covers all learning goals, even though it is randomized.
- Supports inclusion of bitmap images (all file formats that web browsers support, e.g. jpg, png) in the question. The image is shown centre-aligned below the question text.
- Supports LaTeX maths equations, both in the question text and feedback text (when the question is answered wrongly).
- Supports HTML elements, both in question text and feedback. You can therefore format fonts, e.g. &lt;b&gt;bold&lt;/b&gt;, &lt;i&gt;italics&lt;/i&gt;, &lt;u&gt;underline&lt;/u&gt;, line break&lt;br&gt;, &lt;a href..&gt;link&lt;/a&gt;, etc.
- You can exploit automated quiz generation in two alternative ways:
  - You can edit a quiz manually in the spreadsheet, using copy-paste to generate large pools of questions, and optionally exploit Excel features such as cell formulae and referencing, e.g. =RND(), and paramtric text, e.g. ="What is " & ROW() & "×10 ?"
  - ...or you can write VBA code to automatically generate randomized quiz questions! See subroutine '**UserGenerate()**' as an example, which generates 100 questions. You can edit the VBA source code to suit your needs.
- An advantage of quizXL over editing a quiz in CANVAS is that you can have more complex equations (all Excel worksheet functions) in question type Formula Question '**calculated_question**'. CANVAS has a limited set of [helper functions](https://community.canvaslms.com/t5/Canvas-Resource-Documents/Canvas-Formula-Quiz-Question-Helper-Functions/ta-p/387062).
