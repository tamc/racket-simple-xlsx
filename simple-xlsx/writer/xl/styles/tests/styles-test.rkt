#lang racket

(require rackunit/text-ui)

(require "../../../../lib/lib.rkt")

(require rackunit "../styles.rkt")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-styles"
    
    (let ([style_list (list #hash((fill . 1)) #hash((fill . 2)) #hash((fill . 3)))]
          [fill_list (list #hash((fgColor . "FF0000")) #hash((fgColor . "00FF00")) #hash((fgColor . "0000FF")))])
      
      (call-with-input-file "styles-test.dat"
        (lambda (expected)
          (call-with-input-string
           (write-styles style_list fill_list)
           (lambda (actual)
             (check-lines? expected actual)))))
      ))))

(run-tests test-styles)
