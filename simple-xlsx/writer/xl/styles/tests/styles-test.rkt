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
    
    (let ([style_list (list #hash((fill . 1)) #hash((fill . 2)) #hash((fill . 3)))]
          [fill_list (list #hash((backgroundColor . "FF0000")) #hash((backgroundColor . "00FF00")) #hash((backgroundColor . "0000FF")))])
      
      (call-with-input-file test_file
        (lambda (expected)
          (call-with-input-string
           (write-styles style_list fill_list)
           (lambda (actual)
             (check-lines? expected actual)))))
      ))))

(run-tests test-styles)
