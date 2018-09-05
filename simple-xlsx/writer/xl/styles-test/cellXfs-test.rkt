#lang racket

(require rackunit/text-ui)

(require "../../../lib/lib.rkt")

(require rackunit "../styles.rkt")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-cellXfs"
    
    (let ([style_list '()])

      (let ([style_hash (make-hash)])
        (hash-set! style_hash 'fill 1)
        (set! style_list `(,@style_list ,style_hash)))

      (let ([style_hash (make-hash)])
        (hash-set! style_hash 'fill 2)
        (set! style_list `(,@style_list ,style_hash)))

      (let ([style_hash (make-hash)])
        (hash-set! style_hash 'fill 3)
        (set! style_list `(,@style_list ,style_hash)))
      
      (call-with-input-file "cellXfs-test.dat"
        (lambda (expected)
          (call-with-input-string
           (write-cellXfs style_list)
           (lambda (actual)
             (check-lines? expected actual)))))
      ))))

(run-tests test-styles)
