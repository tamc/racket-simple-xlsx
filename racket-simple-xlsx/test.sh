HOME_DIR="racket-simple-xlsx/racket-simple-xlsx"
echo "lib";cd; cd $HOME_DIR/lib/;racket lib-test.rkt
echo "reader-test1";cd; cd $HOME_DIR/reader/tests/test1/;racket test1-test.rkt
echo "reader-test2";cd; cd $HOME_DIR/reader/tests/test2/;racket test2-test.rkt
echo "reader-test3";cd; cd $HOME_DIR/reader/tests/test3/;racket test3-test.rkt
echo "reader-test4";cd; cd $HOME_DIR/reader/tests/test4/;racket test4-test.rkt
echo "reader-test5";cd; cd $HOME_DIR/reader/tests/test5/;racket test5-test.rkt
echo "reader-test6";cd; cd $HOME_DIR/reader/tests/test6/;racket test6-test.rkt

echo "writer-content-type";cd; cd $HOME_DIR/writer/;racket content-type-test.rkt
echo "writer-rels";cd; cd $HOME_DIR/writer/;racket rels-test.rkt
echo "writer-docprops-app";cd; cd $HOME_DIR/writer/;racket docprops-app-test.rkt
