#lang racket

(require rackunit/text-ui)

(require "../../../../lib/lib.rkt")

(require rackunit "../styles.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "styles-test.dat")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-styles"
    
    (let ([style_list 
           (list 
            #hash((fill . 1) (font . 3)) 
            #hash((fill . 2))
            #hash((fill . 3) (font . 1))
            #hash((font . 2))
            )
           ]
          [fill_list 
           (list 
            #hash((fgColor . "FF0000")) 
            #hash((fgColor . "00FF00")) 
            #hash((fgColor . "0000FF")))]
          [font_list 
           (list 
            #hash((fontSize . 20)) 
            #hash((fontSize . 30))
            #hash((fontSize . 40))
            )])

      (call-with-input-file test_file
        (lambda (expected)
          (call-with-input-string
           (write-styles style_list fill_list font_list)
           (lambda (actual)
             (check-lines? expected actual)))))
      ))))

(run-tests test-styles)
