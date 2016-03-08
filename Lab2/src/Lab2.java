import java.util.Iterator;
import java.util.List;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.HashSet;


public class Lab2 {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		String a = "abcddght";
		String b = "fcddgjyt";
		System.out.println(stringIntersect(a, b, 4));
		
		//sameCount check
		List aa = new ArrayList();
		aa.add("a");
		aa.add("y");
		aa.add("a");
		aa.add("b");
		aa.add("y");
		aa.add("b");
		aa.add("d");
		
		List bb = new ArrayList();
		bb.add("o");
		bb.add("a");
		bb.add("a");
		bb.add("o");
		bb.add("b");
		bb.add("b");
		bb.add("l");
		bb.add("d");
		System.out.println(sameCount(aa, bb));
		
		List cc = new ArrayList();
		cc.add("a");
		cc.add("c");
		cc.add("a");
		cc.add("b");
		
		Taboo tb = new Taboo(cc);
		System.out.println(tb.noFollow("a"));
		
		List dd = new ArrayList();
		dd.add("a");
		dd.add("c");
		dd.add("b");
		dd.add("x");
		dd.add("c");
		dd.add("a");
		tb.reduce(dd);
		Iterator it = dd.iterator();
		while(it.hasNext()){
			System.out.print(it.next()+ " ");
		}
	}
    public static boolean stringIntersect(String a, String b, int len){
    	HashSet hs = new HashSet();
    	for(int i=0; i < b.length()-len +1; i++){
    		hs.add(b.substring(i, i+len));
    	}
    	for(int j=0; j< a.length()-len+1;j++){
    		if(hs.contains(a.substring(j, j+len))){
    			return true;
    		}
    	}
    	return false;
    }
    public static <T> int sameCount(Collection<T> a, Collection <T> b){
    	int count =0;
    	for(int i =0; i < a.size();i++){
    		if(Collections.frequency(b, a.toArray()[i]) == Collections.frequency(a, a.toArray()[i])){
    			count++;
    			b.remove(a.toArray()[i]);
    		}
    	}
    	return count;
    }
}
