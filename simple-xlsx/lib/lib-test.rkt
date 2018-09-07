#lang racket

(require rackunit/text-ui)
(require racket/date)

(require rackunit "lib.rkt")

(define test-lib
  (test-suite
   "test-lib"
   
   (test-case
    "test-check-lines1"
    (call-with-input-string 
     "abc"
     (lambda (expected_port)
       (call-with-input-string
        "abc"
        (lambda (test_port)
          (check-lines? expected_port test_port))))))

   (test-case
    "test-check-lines2"
    (call-with-input-string 
     "abc\n11"
     (lambda (expected_port)
       (call-with-input-string
        "abc\n11"
        (lambda (test_port)
          (check-lines? expected_port test_port))))))

   (test-case 
    "test-AZ-NUMBER"
    (check-equal? (abc->number "A") 1)
    (check-equal? (abc->number "B") 2)
    (check-equal? (abc->number "Z") 26)
    (check-equal? (abc->number "AA") 27)
    (check-equal? (abc->number "AB") 28)
    (check-equal? (abc->number "AZ") 52)
    (check-equal? (abc->number "BA") 53)
    (check-equal? (abc->number "YZ") 676)
    (check-equal? (abc->number "ZA") 677)
    (check-equal? (abc->number "ZZ") 702)
    (check-equal? (abc->number "AAA") 703)
   )

   (test-case 
    "test-NUMBER-AZ"
    (check-equal? (number->abc 1) "A")
    (check-equal? (number->abc 26) "Z")
    (check-equal? (number->abc 27) "AA")
    (check-equal? (number->abc 28) "AB")
    (check-equal? (number->abc 29) "AC")
    (check-equal? (number->abc 51) "AY")
    (check-equal? (number->abc 52) "AZ")
    (check-equal? (number->abc 53) "BA")
    (check-equal? (number->abc 676) "YZ")
    (check-equal? (number->abc 677) "ZA")
    (check-equal? (number->abc 702) "ZZ")
    (check-equal? (number->abc 703) "AAA")
    )

   (test-case
    "test-AZ-RANGE"
    (check-equal? (abc->range "A") '(1 . 1))
    (check-equal? (abc->range "B") '(2 . 2))
    (check-equal? (abc->range "2") '(2 . 2))
    (check-equal? (abc->range "C-D") '(3 . 4))
    (check-equal? (abc->range "3-4") '(3 . 4))
    (check-equal? (abc->range "1-26") '(1 . 26))
    (check-equal? (abc->range "26-1") '(1 . 1))
    (check-equal? (abc->range "A-Z") '(1 . 26))
    (check-equal? (abc->range "Z-A") '(1 . 1))
    (check-equal? (abc->range "A-B-C") '(1 . 1))
    (check-equal? (abc->range "A-ksdk344") '(1 . 1))
    (check-equal? (abc->range "sdksjdkf-%^%$#") '(1 . 1))
    )

   (test-case 
    "test-number->list"
    (check-equal? (number->list 1) '(1))
    (check-equal? (number->list 2) '(1 2))
    (check-equal? (number->list 5) '(1 2 3 4 5)))
   
   (test-case
    "test-format-hour"
    (check-equal? (format-time 0.5560763888888889) "13:20:45")
    (check-equal? (format-time 0.2560763888888889) "06:08:45")
    (check-equal? (format-time 0.5) "12:00:00"))
   
   (test-case
    "test-value-of-time"
    (check-equal? (format-complete-time (value-of-time "20141231 12:31:23")) "20141231 12:31:23")
    (check-equal? (format-complete-time (value-of-time "20140105 12:31:23")) "20140105 12:31:23"))

   (test-case
    "test-format-w3cdtf"
    (check-equal? (format-w3cdtf (date* 44 17 13 2 1 2015 5 1 #f 28800 996159076 "CST")) "2015-01-02T13:17:44+08:00"))
   
   (test-case
    "test-create-sheet-name-list"
    (check-equal? (create-sheet-name-list 5) '("Sheet1" "Sheet2" "Sheet3" "Sheet4" "Sheet5"))
    (check-equal? (create-sheet-name-list 1) '("Sheet1")))

   (test-case
    "test-get-dimension"
    (check-equal? (get-dimension '(("1" "2") ("1") () ("1" "2" "3" "4"))) "D4")
    (check-equal? (get-dimension '(("1" "2" "3" "4") ("3" "4" "5" "6"))) "D2")
    )
   
   (test-case
    "test-range-hash-ref"
    
    (let ([range_hash (make-hash)])
      (hash-set! range_hash "A1-A2" 1)
      (hash-set! range_hash "A3-B4" 2)
      (hash-set! range_hash "C1-AA10" 3)
      
      (check-equal? (range-hash-ref range_hash "A1") 1)
      (check-equal? (range-hash-ref range_hash "A2") 1)

      (check-equal? (range-hash-ref range_hash "A3") 2)
      (check-equal? (range-hash-ref range_hash "B3") 2)

      (check-equal? (range-hash-ref range_hash "C8") 3)
      (check-equal? (range-hash-ref range_hash "Z10") 3)
      (check-equal? (range-hash-ref range_hash "AA10") 3)
      ))

   (test-case
    "test-flat-range-hash"
    
    (let ([range_hash (make-hash)]
          [flat_range_hash #f])
      (hash-set! range_hash "A1-A2" 1)
      (hash-set! range_hash "A3-B4" 2)
      
      (set! flat_range_hash (flat-range-hash range_hash))
      
      (check-equal? (hash-count flat_range_hash) 6)
      (check-equal? (hash-ref flat_range_hash "A1") 1)
      (check-equal? (hash-ref flat_range_hash "A2") 1)
      (check-equal? (hash-ref flat_range_hash "A3") 2)
      (check-equal? (hash-ref flat_range_hash "A4") 2)
      (check-equal? (hash-ref flat_range_hash "B3") 2)
      (check-equal? (hash-ref flat_range_hash "B4") 2))

    (let ([range_hash (make-hash)]
          [flat_range_hash #f])
      (hash-set! range_hash "A2-A1" 1)
      (hash-set! range_hash "A5-E14" 2)
      
      (set! flat_range_hash (flat-range-hash range_hash))
      
      (check-equal? (hash-count flat_range_hash) 50)
      (check-equal? (hash-ref flat_range_hash "A5") 2)
      (check-equal? (hash-ref flat_range_hash "B10") 2)
      (check-equal? (hash-ref flat_range_hash "E14") 2)
      )
      
    )

   (test-case
    "test-prefix-each-line"
    
    (let ([str "kkd\nskdfk\n\nksjdkf\n\n"])
      (check-equal? (prefix-each-line str "  ")
                    "  kkd\n  skdfk\n  \n  ksjdkf\n  \n"))

    (let ([str "kkd\nskdfk\n\nksjdkf"])
      (check-equal? (prefix-each-line str "  ")
                    "  kkd\n  skdfk\n  \n  ksjdkf\n"))

    (let ([str "kkd\nskdfk\n\nksjdkf\n"])
      (check-equal? (prefix-each-line str "  ")
                    "  kkd\n  skdfk\n  \n  ksjdkf\n"))

    (let ([str ""])
      (check-equal? (prefix-each-line str "  ") "  \n"))

    (let ([str "\n"])
      (check-equal? (prefix-each-line str "  ") "  \n"))
    )
    
  ))

(run-tests test-lib)
