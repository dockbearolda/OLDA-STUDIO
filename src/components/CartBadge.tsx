import { motion, AnimatePresence } from 'framer-motion';
import { useCart } from '../context/CartContext';

interface Props {
  onClick: () => void;
}

export default function CartBadge({ onClick }: Props) {
  const { items } = useCart();
  const count = items.length;

  return (
    <button className="cart-badge-btn" onClick={onClick} aria-label="Panier">
      {/* Shopping bag icon */}
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
        <path d="M6 2L3 6v14a2 2 0 002 2h14a2 2 0 002-2V6l-3-4z"/>
        <line x1="3" y1="6" x2="21" y2="6"/>
        <path d="M16 10a4 4 0 01-8 0"/>
      </svg>

      <AnimatePresence>
        {count > 0 && (
          <motion.span
            className="cart-count"
            initial={{ scale: 0, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            exit={{ scale: 0, opacity: 0 }}
            transition={{ type: 'spring', stiffness: 500, damping: 30 }}
          >
            {count}
          </motion.span>
        )}
      </AnimatePresence>
    </button>
  );
}
