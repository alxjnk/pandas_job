#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Aug  6 14:13:31 2017

@author: alex
"""

import urllib.request
import urllib

from bs4 import BeautifulSoup

import json


BASE_URLS = 'http://www.internetdoc.ru/image/tid/59'


def get_html(url):
    response = urllib.request.urlopen(url)
    return response.read()
    
    
def parse(html):
    soup = BeautifulSoup(html, 'lxml')
    img_list = soup.find_all("span", class_="image-gallery-view-cover-thumbnail")    
    imgs = {img_list[i].img['alt']: img_list[i].img['src']  
        for i in range(48)} #write data to dict  
    print(img_list[1].img['src'])
    print(imgs)            
    return imgs
    
         
def saver(data, path):
    with open(path, 'w') as fp:
        json.dump(data, fp)
    
           
       
def main():
     data = parse(get_html(BASE_URLS))
    
    
    
if __name__ == '__main__':
    main()