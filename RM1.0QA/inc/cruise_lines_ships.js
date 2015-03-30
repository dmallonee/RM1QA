
 
var arrayData = new Array(); 
var i = -1


arrayData[i++]	= 'Carnival|Carnival Conquest|'
arrayData[i++]	= 'Carnival|Carnival Destiny|'
arrayData[i++]	= 'Carnival|Carnival Glory|'
arrayData[i++]	= 'Carnival|Carnival Legend|'
arrayData[i++]	= 'Carnival|Carnival Miracle|'
arrayData[i++]	= 'Carnival|Carnival Pride|'
arrayData[i++]	= 'Carnival|Carnival Spirit|'
arrayData[i++]	= 'Carnival|Carnival Triumph|'
arrayData[i++]	= 'Carnival|Carnival Victory|'
arrayData[i++]	= 'Carnival|Celebration|'
arrayData[i++]	= 'Carnival|Ecstasy|'
arrayData[i++]	= 'Carnival|Elation|'
arrayData[i++]	= 'Carnival|Fantasy|'
arrayData[i++]	= 'Carnival|Fascination|'
arrayData[i++]	= 'Carnival|Holiday|'
arrayData[i++]	= 'Carnival|Imagination|'
arrayData[i++]	= 'Carnival|Inspiration|'
arrayData[i++]	= 'Carnival|Jubilee|'
arrayData[i++]	= 'Carnival|Paradise|'
arrayData[i++]	= 'Carnival|Sensation|'

arrayData[i++]	= 'Celebrity|Century|'
arrayData[i++]	= 'Celebrity|Constellation|'
arrayData[i++]	= 'Celebrity|Galaxy|'
arrayData[i++]	= 'Celebrity|Horizon|'
arrayData[i++]	= 'Celebrity|Infinity|'
arrayData[i++]	= 'Celebrity|Mercury|'
arrayData[i++]	= 'Celebrity|Millennium|'
arrayData[i++]	= 'Celebrity|Summit|'
arrayData[i++]	= 'Celebrity|Zenith|'

arrayData[i++]	= 'Crystal|Crystal Harmony|'
arrayData[i++]	= 'Crystal|Crystal Serenity|'
arrayData[i++]	= 'Crystal|Crystal Symphony|'

arrayData[i++]	= 'Disney|Disney Magic|'
arrayData[i++]	= 'Disney|Disney Wonder|'

arrayData[i++]	= 'Holland America|Amsterdam|'
arrayData[i++]	= 'Holland America|Maasdam|'
arrayData[i++]	= 'Holland America|Noordam|'
arrayData[i++]	= 'Holland America|Oosterdam|'
arrayData[i++]	= 'Holland America|Prinsendam|'
arrayData[i++]	= 'Holland America|Rotterdam|'
arrayData[i++]	= 'Holland America|Ryndam|'
arrayData[i++]	= 'Holland America|Statendam|'
arrayData[i++]	= 'Holland America|Veendam|'
arrayData[i++]	= 'Holland America|Volendam|'
arrayData[i++]	= 'Holland America|Westerdam|'
arrayData[i++]	= 'Holland America|Zaandam|'
arrayData[i++]	= 'Holland America|Zuiderdam|'

arrayData[i++]	= 'Norwegian|Norway|'
arrayData[i++]	= 'Norwegian|Norwegian Crown|'
arrayData[i++]	= 'Norwegian|Norwegian Dawn|'
arrayData[i++]	= 'Norwegian|Norwegian Dream|'
arrayData[i++]	= 'Norwegian|Norwegian Majesty|'
arrayData[i++]	= 'Norwegian|Norwegian Sea|'
arrayData[i++]	= 'Norwegian|Norwegian Sky|'
arrayData[i++]	= 'Norwegian|Norwegian Star|'
arrayData[i++]	= 'Norwegian|Norwegian Sun|'
arrayData[i++]	= 'Norwegian|Norwegian Wind|'
arrayData[i++]	= 'Norwegian|Pride of Aloha|'
arrayData[i++]	= 'Norwegian|Pride of America|'

arrayData[i++]	= 'Princess|Caribbean Princess|'
arrayData[i++]	= 'Princess|Coral Princess|'
arrayData[i++]	= 'Princess|Dawn Princess|'
arrayData[i++]	= 'Princess|Diamond Princess|'
arrayData[i++]	= 'Princess|Golden Princess|'
arrayData[i++]	= 'Princess|Grand Princess|'
arrayData[i++]	= 'Princess|Island Princess|'
arrayData[i++]	= 'Princess|Pacific Princess|'
arrayData[i++]	= 'Princess|Regal Princess|'
arrayData[i++]	= 'Princess|Royal Princess|'
arrayData[i++]	= 'Princess|Sapphire Princess|'
arrayData[i++]	= 'Princess|Star Princess|'
arrayData[i++]	= 'Princess|Sun Princess|'
arrayData[i++]	= 'Princess|Tahitian Princess|'

arrayData[i++]	= 'Royal Caribbean|Adventure of the Seas|'
arrayData[i++]	= 'Royal Caribbean|Brilliance of the Seas|'
arrayData[i++]	= 'Royal Caribbean|Enchantment of the Seas|'
arrayData[i++]	= 'Royal Caribbean|Explorer of the Seas|'
arrayData[i++]	= 'Royal Caribbean|Grandeur of the Seas|'
arrayData[i++]	= 'Royal Caribbean|Jewel of the Seas|'
arrayData[i++]	= 'Royal Caribbean|Legend of the Seas|'
arrayData[i++]	= 'Royal Caribbean|Majesty of the Seas|'
arrayData[i++]	= 'Royal Caribbean|Mariner of the Seas|'
arrayData[i++]	= 'Royal Caribbean|Monarch of the Seas|'
arrayData[i++]	= 'Royal Caribbean|Navigator of the Seas|'
arrayData[i++]	= 'Royal Caribbean|Nordic Empress|'
arrayData[i++]	= 'Royal Caribbean|Radiance of the Seas|'
arrayData[i++]	= 'Royal Caribbean|Rhapsody of the Seas|'
arrayData[i++]	= 'Royal Caribbean|Serenade of the Seas|'
arrayData[i++]	= 'Royal Caribbean|Sovereign of the Seas|'
arrayData[i++]	= 'Royal Caribbean|Splendour of the Seas|'
arrayData[i++]	= 'Royal Caribbean|Vision of the Seas|'
arrayData[i++]	= 'Royal Caribbean|Voyager of the Seas|'

arrayData[i++]	= 'Windstar|Wind Spirit|'
arrayData[i++]	= 'Windstar|Wind Star|'
arrayData[i++]	= 'Windstar|Wind Surf|'


function PopulateShipData( name ) { 
 
	select	= window.document.search_criteria.selected_ships; 
	string	= ""; 
 
		// 0 - will display the new options only 
		// 1 - will display the first existing option plus the new options 
 
	count	= 0; 
 
		// Clear the old list (above element 0) 
 
	select.options.length = count; 
 
		// Place all matching categories into Options. 
 
	for( i = 0; i < arrayData.length; i++ ) { 
		string = arrayData[i].split( "|" ); 
		if( string[0] == name ) { 
			select.options[count++] = new Option( string[1] ); 
		} 
	} 
 
		// Set which option from subcategory is to be selected 
 
//	select.options.selectedIndex = 2; 
 
		// Give subcategory focus and select it 
 
//	select.focus(); 
 
} 
 

