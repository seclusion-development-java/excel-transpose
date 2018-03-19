
public abstract class Search {

	public SearchCible searchCible;

	public Search() {

	}

	public abstract void afficher();

	public void effectuerSearch() {
		searchCible.search();
	}



	public void nager() {
	 System.out.println("Tous les canards flottent, même les leurres!");
	 }
}
