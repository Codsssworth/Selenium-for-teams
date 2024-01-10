package endter9003;
import java.util.*;
public class ShortSentences {
	static Scanner s = new Scanner(System.in);
	
	public static void main(String[] args) {
		// Taking input
		System.out.println("Input a sentence");
		String completeSentence = s.nextLine();
		
		// Breaking the string to an array of strings and add to set to make it unique by itself
		String[] individualWords = completeSentence.strip().split(" ");
		Set<String> unique = new HashSet<>();
		for(String st : individualWords) {
			unique.add(st);
		}
		
		// Add all the unique words to a new array list for sorting
		ArrayList<String> uniqueWords = new ArrayList<>();
		Iterator<String> it = unique.iterator();
		while(it.hasNext()) {
			// Changing the word to lower case for sorting all the words properly
			uniqueWords.add(it.next().toLowerCase());
		}
		
		// Sorting the array
		Collections.sort(uniqueWords);
		
		// Creating and adding every word to a linked list
		LinkedList<String> sortedWords = new LinkedList<>();
		for(String st : uniqueWords) {
			sortedWords.add(st);
		}
		
		// Print all the words from the linked list using iterator
		Iterator<String> itr = sortedWords.iterator();
		while(itr.hasNext()) {
			System.out.println(itr.next());
		}
	}
}


