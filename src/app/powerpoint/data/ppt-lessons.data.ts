import { ActivityPayload } from '../../shared/models/activity.types';

export const PPT_SAMPLE_LESSONS: Record<string, ActivityPayload> = {
  "slideTitle": {
    "app": "powerpoint",
    "mode": "lesson",
    "title": "Add a Slide Title",
    "assetBaseUrl": "./assets/",
    "theme": "LIGHT",
    "images": [
      {
        "file": "headphones.svg",
        "alt": "Headphones"
      },
      {
        "file": "books-stack.svg",
        "alt": "A stack of books"
      },
      {
        "file": "art.svg",
        "alt": "Paint and brushes"
      },
      {
        "file": "sports.svg",
        "alt": "Sports balls"
      }
    ],
    "tasks": [
      {
        "title": "Add the title",
        "explanation": "A title tells what a slide is about. It usually goes at the top.",
        "instructions": [
          "Click the title area on the slide and type `My Favorite Animal`."
        ],
        "hint": "Click inside the box that says Click to add title, then type.",
        "startingDeck": {
          "theme": "simple-light",
          "slides": [
            {
              "layout": "titleAndBody",
              "title": "",
              "body": ""
            }
          ]
        },
        "goals": [
          {
            "type": "slideText",
            "slide": 1,
            "placeholder": "title",
            "contains": "My Favorite Animal"
          }
        ],
        "showMeHow": true
      }
    ]
  },
  "newSlide": {
    "app": "powerpoint",
    "mode": "lesson",
    "title": "Add a New Slide",
    "assetBaseUrl": "./assets/",
    "theme": "LIGHT",
    "images": [
      {
        "file": "headphones.svg",
        "alt": "Headphones"
      },
      {
        "file": "books-stack.svg",
        "alt": "A stack of books"
      },
      {
        "file": "art.svg",
        "alt": "Paint and brushes"
      },
      {
        "file": "sports.svg",
        "alt": "Sports balls"
      }
    ],
    "tasks": [
      {
        "title": "Add a new slide",
        "explanation": "Presentations usually have several slides to organize ideas.",
        "instructions": [
          "Click Home > New Slide to add a slide."
        ],
        "hint": "Click the New Slide button on the ribbon.",
        "startingDeck": {
          "theme": "simple-light",
          "slides": [
            {
              "layout": "title",
              "title": "My Presentation",
              "subtitle": ""
            }
          ]
        },
        "goals": [
          {
            "type": "slideCount",
            "min": 2
          }
        ],
        "showMeHow": true
      }
    ]
  },
  "changeLayout": {
    "app": "powerpoint",
    "mode": "lesson",
    "title": "Change the Layout",
    "assetBaseUrl": "./assets/",
    "theme": "LIGHT",
    "images": [
      {
        "file": "headphones.svg",
        "alt": "Headphones"
      },
      {
        "file": "books-stack.svg",
        "alt": "A stack of books"
      },
      {
        "file": "art.svg",
        "alt": "Paint and brushes"
      },
      {
        "file": "sports.svg",
        "alt": "Sports balls"
      }
    ],
    "tasks": [
      {
        "title": "Change slide layout",
        "explanation": "Layouts arrange placeholders on a slide for different kinds of content.",
        "instructions": [
          "Click Home > Layout and choose `Two Content`."
        ],
        "hint": "Click the Layout button on the ribbon and select Two Content.",
        "startingDeck": {
          "theme": "simple-light",
          "slides": [
            {
              "layout": "titleAndBody",
              "title": "Comparing Ideas",
              "body": "First point"
            }
          ]
        },
        "goals": [
          {
            "type": "layout",
            "slide": 1,
            "layout": "twoColumns"
          }
        ],
        "showMeHow": true
      }
    ]
  },
  "textStyle": {
    "app": "powerpoint",
    "mode": "lesson",
    "title": "Format Slide Text",
    "assetBaseUrl": "./assets/",
    "theme": "LIGHT",
    "images": [
      {
        "file": "headphones.svg",
        "alt": "Headphones"
      },
      {
        "file": "books-stack.svg",
        "alt": "A stack of books"
      },
      {
        "file": "art.svg",
        "alt": "Paint and brushes"
      },
      {
        "file": "sports.svg",
        "alt": "Sports balls"
      }
    ],
    "tasks": [
      {
        "title": "Format the title text",
        "explanation": "Changing font size and weight emphasizes important headings.",
        "instructions": [
          "Select the title text `All about me`, make it bold, and change the font size to `54`."
        ],
        "hint": "Select the title text, then click Home > Bold and set the font size to 54.",
        "startingDeck": {
          "theme": "simple-light",
          "slides": [
            {
              "layout": "title",
              "title": "All about me",
              "subtitle": ""
            }
          ]
        },
        "goals": [
          {
            "type": "textStyle",
            "slide": 1,
            "text": "All about me",
            "bold": true,
            "size": 54
          }
        ],
        "showMeHow": true
      }
    ]
  },
  "list": {
    "app": "powerpoint",
    "mode": "lesson",
    "title": "Add a Bulleted List",
    "assetBaseUrl": "./assets/",
    "theme": "LIGHT",
    "images": [
      {
        "file": "headphones.svg",
        "alt": "Headphones"
      },
      {
        "file": "books-stack.svg",
        "alt": "A stack of books"
      },
      {
        "file": "art.svg",
        "alt": "Paint and brushes"
      },
      {
        "file": "sports.svg",
        "alt": "Sports balls"
      }
    ],
    "tasks": [
      {
        "title": "Create a bulleted list",
        "explanation": "Bulleted lists help break down key information into readable points.",
        "instructions": [
          "Click the body placeholder and type 3 bullet items."
        ],
        "hint": "Click inside the content box and enter at least 3 bulleted lines.",
        "startingDeck": {
          "theme": "simple-light",
          "slides": [
            {
              "layout": "titleAndBody",
              "title": "My Goals",
              "body": ""
            }
          ]
        },
        "goals": [
          {
            "type": "list",
            "slide": 1,
            "placeholder": "body",
            "listType": "bullet",
            "minItems": 3
          }
        ],
        "showMeHow": true
      }
    ]
  },
  "image": {
    "app": "powerpoint",
    "mode": "lesson",
    "title": "Insert a Picture",
    "assetBaseUrl": "./assets/",
    "theme": "LIGHT",
    "images": [
      {
        "file": "headphones.svg",
        "alt": "Headphones"
      },
      {
        "file": "books-stack.svg",
        "alt": "A stack of books"
      },
      {
        "file": "art.svg",
        "alt": "Paint and brushes"
      },
      {
        "file": "sports.svg",
        "alt": "Sports balls"
      }
    ],
    "tasks": [
      {
        "title": "Add a picture to the slide",
        "explanation": "Pictures make slides more visually engaging.",
        "instructions": [
          "Click Insert > Pictures and choose any picture."
        ],
        "hint": "Click Insert > Pictures to open the picture gallery and select an image.",
        "startingDeck": {
          "theme": "simple-light",
          "slides": [
            {
              "layout": "titleAndBody",
              "title": "Hobby Showcase",
              "body": ""
            }
          ]
        },
        "goals": [
          {
            "type": "image",
            "slide": 1,
            "file": null
          }
        ],
        "showMeHow": true
      }
    ]
  },
  "textBox": {
    "app": "powerpoint",
    "mode": "lesson",
    "title": "Insert a Text Box",
    "assetBaseUrl": "./assets/",
    "theme": "LIGHT",
    "images": [
      {
        "file": "headphones.svg",
        "alt": "Headphones"
      },
      {
        "file": "books-stack.svg",
        "alt": "A stack of books"
      },
      {
        "file": "art.svg",
        "alt": "Paint and brushes"
      },
      {
        "file": "sports.svg",
        "alt": "Sports balls"
      }
    ],
    "tasks": [
      {
        "title": "Add a custom text box",
        "explanation": "Text boxes allow you to place text anywhere on a slide.",
        "instructions": [
          "Click Insert > Text Box and type `Source: class survey`."
        ],
        "hint": "Click Insert > Text Box, click on the slide, and type Source: class survey.",
        "startingDeck": {
          "theme": "simple-light",
          "slides": [
            {
              "layout": "titleAndBody",
              "title": "Survey Results",
              "body": "80% of students voted yes."
            }
          ]
        },
        "goals": [
          {
            "type": "textBox",
            "slide": 1,
            "contains": "Source: class survey"
          }
        ],
        "showMeHow": true
      }
    ]
  },
  "shape": {
    "app": "powerpoint",
    "mode": "lesson",
    "title": "Insert a Shape",
    "assetBaseUrl": "./assets/",
    "theme": "LIGHT",
    "images": [
      {
        "file": "headphones.svg",
        "alt": "Headphones"
      },
      {
        "file": "books-stack.svg",
        "alt": "A stack of books"
      },
      {
        "file": "art.svg",
        "alt": "Paint and brushes"
      },
      {
        "file": "sports.svg",
        "alt": "Sports balls"
      }
    ],
    "tasks": [
      {
        "title": "Add a star shape",
        "explanation": "Shapes can highlight key elements or decorate slides.",
        "instructions": [
          "Click Insert > Shapes and choose a star shape."
        ],
        "hint": "Click Insert > Shapes and select the star.",
        "startingDeck": {
          "theme": "simple-light",
          "slides": [
            {
              "layout": "blank",
              "elements": []
            }
          ]
        },
        "goals": [
          {
            "type": "shape",
            "slide": 1,
            "shape": "star"
          }
        ],
        "showMeHow": true
      }
    ]
  },
  "table": {
    "app": "powerpoint",
    "mode": "lesson",
    "title": "Insert a Table",
    "assetBaseUrl": "./assets/",
    "theme": "LIGHT",
    "images": [
      {
        "file": "headphones.svg",
        "alt": "Headphones"
      },
      {
        "file": "books-stack.svg",
        "alt": "A stack of books"
      },
      {
        "file": "art.svg",
        "alt": "Paint and brushes"
      },
      {
        "file": "sports.svg",
        "alt": "Sports balls"
      }
    ],
    "tasks": [
      {
        "title": "Add a data table",
        "explanation": "Tables organize information into rows and columns.",
        "instructions": [
          "Click Insert > Table and create a table with 3 rows and 2 columns."
        ],
        "hint": "Click Insert > Table and choose 2 columns by 3 rows.",
        "startingDeck": {
          "theme": "simple-light",
          "slides": [
            {
              "layout": "titleOnly",
              "title": "Class Schedule",
              "elements": []
            }
          ]
        },
        "goals": [
          {
            "type": "table",
            "slide": 1,
            "rows": 3,
            "columns": 2
          }
        ],
        "showMeHow": true
      }
    ]
  },
  "background": {
    "app": "powerpoint",
    "mode": "lesson",
    "title": "Change Background Color",
    "assetBaseUrl": "./assets/",
    "theme": "LIGHT",
    "images": [
      {
        "file": "headphones.svg",
        "alt": "Headphones"
      },
      {
        "file": "books-stack.svg",
        "alt": "A stack of books"
      },
      {
        "file": "art.svg",
        "alt": "Paint and brushes"
      },
      {
        "file": "sports.svg",
        "alt": "Sports balls"
      }
    ],
    "tasks": [
      {
        "title": "Change slide background",
        "explanation": "Custom background colors change the visual style of your presentation.",
        "instructions": [
          "Click Design > Format Background and choose `Light Blue`."
        ],
        "hint": "Click Design > Format Background and pick the Light Blue color swatch.",
        "startingDeck": {
          "theme": "simple-light",
          "slides": [
            {
              "layout": "title",
              "title": "Ocean Life",
              "subtitle": ""
            }
          ]
        },
        "goals": [
          {
            "type": "background",
            "slide": 1,
            "color": "lightBlue"
          }
        ],
        "showMeHow": true
      }
    ]
  },
  "theme": {
    "app": "powerpoint",
    "mode": "lesson",
    "title": "Apply a Theme",
    "assetBaseUrl": "./assets/",
    "theme": "LIGHT",
    "images": [
      {
        "file": "headphones.svg",
        "alt": "Headphones"
      },
      {
        "file": "books-stack.svg",
        "alt": "A stack of books"
      },
      {
        "file": "art.svg",
        "alt": "Paint and brushes"
      },
      {
        "file": "sports.svg",
        "alt": "Sports balls"
      }
    ],
    "tasks": [
      {
        "title": "Change the presentation theme",
        "explanation": "Themes coordinate colors, fonts, and layouts across your slides.",
        "instructions": [
          "Click Design and choose any theme other than the default theme."
        ],
        "hint": "Open the Design tab and click any theme tile in the Themes gallery.",
        "startingDeck": {
          "theme": "simple-light",
          "slides": [
            {
              "layout": "title",
              "title": "Science Fair",
              "subtitle": "Grade 5"
            }
          ]
        },
        "goals": [
          {
            "type": "theme",
            "theme": "any"
          }
        ],
        "showMeHow": true
      }
    ]
  },
  "transition": {
    "app": "powerpoint",
    "mode": "lesson",
    "title": "Add a Slide Transition",
    "assetBaseUrl": "./assets/",
    "theme": "LIGHT",
    "images": [
      {
        "file": "headphones.svg",
        "alt": "Headphones"
      },
      {
        "file": "books-stack.svg",
        "alt": "A stack of books"
      },
      {
        "file": "art.svg",
        "alt": "Paint and brushes"
      },
      {
        "file": "sports.svg",
        "alt": "Sports balls"
      }
    ],
    "tasks": [
      {
        "title": "Apply a slide transition",
        "explanation": "Transitions animate how slides appear when presenting.",
        "instructions": [
          "Click Transitions and select the `Fade` transition."
        ],
        "hint": "Open the Transitions tab and click the Fade transition tile.",
        "startingDeck": {
          "theme": "simple-light",
          "slides": [
            {
              "layout": "title",
              "title": "Welcome",
              "subtitle": ""
            }
          ]
        },
        "goals": [
          {
            "type": "transition",
            "slide": 1,
            "transition": "fade"
          }
        ],
        "showMeHow": true
      }
    ]
  },
  "slideOrder": {
    "app": "powerpoint",
    "mode": "lesson",
    "title": "Reorder Slides",
    "assetBaseUrl": "./assets/",
    "theme": "LIGHT",
    "images": [
      {
        "file": "headphones.svg",
        "alt": "Headphones"
      },
      {
        "file": "books-stack.svg",
        "alt": "A stack of books"
      },
      {
        "file": "art.svg",
        "alt": "Paint and brushes"
      },
      {
        "file": "sports.svg",
        "alt": "Sports balls"
      }
    ],
    "tasks": [
      {
        "title": "Arrange slides in logical order",
        "explanation": "Slide order determines the presentation flow.",
        "instructions": [
          "Reorder the slides so they appear in this order: `Introduction`, `The Problem`, `Our Solution`."
        ],
        "hint": "Arrange the slides in order: Introduction, The Problem, then Our Solution.",
        "startingDeck": {
          "theme": "simple-light",
          "slides": [
            {
              "layout": "title",
              "title": "Our Solution"
            },
            {
              "layout": "title",
              "title": "Introduction"
            },
            {
              "layout": "title",
              "title": "The Problem"
            }
          ]
        },
        "goals": [
          {
            "type": "slideOrder",
            "titles": [
              "Introduction",
              "The Problem",
              "Our Solution"
            ]
          }
        ],
        "showMeHow": true
      }
    ]
  },
  "action": {
    "app": "powerpoint",
    "mode": "lesson",
    "title": "Select a Slide",
    "assetBaseUrl": "./assets/",
    "theme": "LIGHT",
    "images": [
      {
        "file": "headphones.svg",
        "alt": "Headphones"
      },
      {
        "file": "books-stack.svg",
        "alt": "A stack of books"
      },
      {
        "file": "art.svg",
        "alt": "Paint and brushes"
      },
      {
        "file": "sports.svg",
        "alt": "Sports balls"
      }
    ],
    "tasks": [
      {
        "title": "Switch slides in the thumbnail panel",
        "explanation": "Clicking a slide thumbnail selects that slide for editing.",
        "instructions": [
          "Click the slide 2 thumbnail in the left panel to select it."
        ],
        "hint": "Click on slide 2 in the thumbnail panel on the left.",
        "startingDeck": {
          "theme": "simple-light",
          "slides": [
            {
              "layout": "title",
              "title": "Slide 1"
            },
            {
              "layout": "titleAndBody",
              "title": "Slide 2",
              "body": "Second slide content"
            }
          ]
        },
        "goals": [
          {
            "type": "action",
            "action": "selectSlide",
            "slide": 2
          }
        ],
        "showMeHow": true
      }
    ]
  },
  "assessment": {
    "app": "powerpoint",
    "mode": "assessment",
    "title": "Presentation Basics Assessment",
    "assetBaseUrl": "./assets/",
    "theme": "LIGHT",
    "images": [
      {
        "file": "headphones.svg",
        "alt": "Headphones"
      },
      {
        "file": "books-stack.svg",
        "alt": "A stack of books"
      },
      {
        "file": "art.svg",
        "alt": "Paint and brushes"
      },
      {
        "file": "sports.svg",
        "alt": "Sports balls"
      }
    ],
    "tasks": [
      {
        "title": "Add slide title",
        "explanation": "Enter a title for the first slide.",
        "instructions": [
          "Click the title box and type `Project Overview`."
        ],
        "hint": "Type Project Overview into the title placeholder.",
        "startingDeck": {
          "theme": "simple-light",
          "slides": [
            {
              "layout": "title",
              "title": "",
              "subtitle": ""
            }
          ]
        },
        "goals": [
          {
            "type": "slideText",
            "slide": 1,
            "placeholder": "title",
            "contains": "Project Overview"
          }
        ]
      },
      {
        "title": "Add a second slide",
        "explanation": "Add another slide to your deck.",
        "instructions": [
          "Click Home > New Slide to add a new slide."
        ],
        "hint": "Click New Slide on the Home tab.",
        "startingDeck": null,
        "goals": [
          {
            "type": "slideCount",
            "min": 2
          }
        ]
      },
      {
        "title": "Change second slide layout",
        "explanation": "Set the layout for the new slide.",
        "instructions": [
          "Change slide 2 to the `Two Content` layout."
        ],
        "hint": "Select slide 2, then click Home > Layout and choose Two Content.",
        "startingDeck": null,
        "goals": [
          {
            "type": "layout",
            "slide": 2,
            "layout": "twoColumns"
          }
        ]
      }
    ]
  }
};
