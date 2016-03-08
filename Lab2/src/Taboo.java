import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.Iterator;

public class Taboo<T> {
	private List<T> Rules;
	public Taboo(List<T> t){
		Rules = t;
	}
	public Set<T> noFollow(T element){
		Set <T> set = new HashSet<T>();
		Iterator<T> it = Rules.iterator();
		T current = null;
		T next = null;
		while(it.hasNext()){
			current = it.next();
			if(current != null && current.equals(element) && it.hasNext()){
				next = it.next();
				if( next != null){
					set.add(next);
				}
			}
		}
		return set;
	}
	public void reduce(List<T> list){
		int flag = -1;
		for(int i=0; i <list.size()-1; i++){
			T current = list.get(i);
			T next = list.get(i+1);
			if( noFollow(current).contains(next)){
				flag =i;
				break;
			}
		}
		if(flag != -1){
			list.remove(flag+1);
			reduce(list);
		}
	}
}
