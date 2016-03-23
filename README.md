### imgtoxlsx

Small ruby script to convert images to the only acceptable format - xlsx.
1px -> 1 cell.

#### Dependencies

imagemagick

```
brew install imagemagick
```


or whatever you use


gems:

```
gem install axlsx rmagick progress_bar
```



#### Usage

```
./imgtoxlsx.rb -i 1.png,2.png,3.png -o output_file.xlsx
```

(each image is a separate sheet in output file)