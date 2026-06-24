export function classifyPaginationDelta(delta) {
	if (delta > 0) {
		return 'over-pagination';
	}
	if (delta < 0) {
		return 'reference-difference';
	}
	return 'match';
}

export function formatPaginationAuditStatus(classification) {
	switch (classification) {
		case 'over-pagination':
			return 'OVER';
		case 'reference-difference':
			return 'REFERENCE-DIFF';
		default:
			return 'PASS';
	}
}
