package edu.handong.csee.java3;

import java.util.LinkedList;

//Customized Generics
public class GenQueue<String> {
	 private LinkedList<String> list = new LinkedList<String>();
	   public void enqueue(String item) {
	      list.addLast(item);
	   }
	   public String dequeue() {
	      return list.poll();
	   }
	   public boolean hasItems() {
	      return !list.isEmpty();
	   }
	   public int size() {
	      return list.size();
	   }
	   public void addItems(GenQueue<? extends String> q) {
	      while (q.hasItems()) list.addLast(q.dequeue());
	   }
}
